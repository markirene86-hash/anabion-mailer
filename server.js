const express = require('express');
const juice = require('juice');

// ─── CSS VARIABLE RESOLVER ────────────────────────────────────────────────────
// Outlook does not support CSS variables (var(--x)) at all.
// This function extracts :root variable definitions and replaces all var() usages.
function resolveCssVars(html) {
  // Extract all :root variables from <style> blocks
  const vars = {};
  
  // Find :root block
  const rootMatch = html.match(/:root\s*\{([^}]+)\}/);
  if (rootMatch) {
    const rootBlock = rootMatch[1];
    // Parse each variable
    const varRx = /--([\w-]+)\s*:\s*([^;]+);/g;
    let m;
    while ((m = varRx.exec(rootBlock)) !== null) {
      vars['--' + m[1].trim()] = m[2].trim();
    }
  }

  if (Object.keys(vars).length === 0) return html;

  // Resolve nested vars (vars that reference other vars)
  function resolveValue(val, depth = 0) {
    if (depth > 5) return val;
    return val.replace(/var\(([^,)]+)(?:,\s*([^)]+))?\)/g, (match, varName, fallback) => {
      varName = varName.trim();
      if (vars[varName]) return resolveValue(vars[varName], depth + 1);
      if (fallback) return resolveValue(fallback.trim(), depth + 1);
      return match;
    });
  }

  // Resolve all var() values
  for (const key in vars) {
    vars[key] = resolveValue(vars[key]);
  }

  // Replace all var() occurrences in the full HTML
  let resolved = html.replace(/var\(([^,)]+)(?:,\s*([^)]+))?\)/g, (match, varName, fallback) => {
    varName = varName.trim();
    if (vars[varName]) return vars[varName];
    if (fallback) return fallback.trim();
    return match;
  });

  // Second pass for any remaining nested vars
  resolved = resolved.replace(/var\(([^,)]+)(?:,\s*([^)]+))?\)/g, (match, varName, fallback) => {
    varName = varName.trim();
    if (vars[varName]) return vars[varName];
    if (fallback) return fallback.trim();
    return match;
  });

  return resolved;
}


const nodemailer = require('nodemailer');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '15mb' }));
app.use(express.urlencoded({ extended: true, limit: '15mb' }));
app.use(express.static(path.join(__dirname, '.')));

// ─── SMTP CONFIG ─────────────────────────────────────────────────────────────
const SMTP = {
  host: 'smtp.office365.com',
  port: 587,
  secure: false,
  auth: {
    user: 'svc.travel-expense-noreply@anabion.com',
    pass: '<HG20#?9MV'
  },
  tls: { rejectUnauthorized: false, minVersion: 'TLSv1.2' },
  requireTLS: true
};

const FROM = process.env.SMTP_FROM || 'Travel & Expense <travel-expense-noreply@anabion.com>';

function createTransport() {
  return nodemailer.createTransport(SMTP);
}

// ─── CSS INLINER ─────────────────────────────────────────────────────────────
function inlineCss(html) {
  try {
    // Step 1: resolve CSS variables (var(--x)) → real values
    // Outlook does not support CSS variables at all
    const resolved = resolveCssVars(html);

    // Step 2: inline all CSS with juice
    const inlined = juice(resolved, {
      removeStyleTags: true,
      preserveMediaQueries: false,
      preserveFontFaces: false,
      applyWidthAttributes: true,
      applyHeightAttributes: true,
      preserveImportant: true,
      extraCss: ''
    });

    return inlined;
  } catch(e) {
    console.error('inlineCss error:', e.message);
    return html;
  }
}

function parseCssRules(css) {
  const rules = [];
  let clean = css
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/@import[^;]*;/g, '')
    .replace(/@charset[^;]*;/g, '')
    .replace(/@(?:font-face|keyframes)[^{]*\{[^{}]*\}/g, '')
    .replace(/@media[^{]*\{[\s\S]*?\}\s*\}/g, '');

  const rx = /([^{}@][^{}]*?)\s*\{\s*([^{}]+?)\s*\}/g;
  let m;
  while ((m = rx.exec(clean)) !== null) {
    const selector = m[1].trim();
    const declarations = m[2].trim();
    if (!selector || !declarations || selector.startsWith('@')) continue;

    const props = {};
    declarations.split(';').forEach(d => {
      const i = d.indexOf(':');
      if (i > 0) {
        const k = d.slice(0, i).trim().toLowerCase();
        const v = d.slice(i + 1).trim();
        if (k && v) props[k] = v;
      }
    });
    rules.push({ selector, props });
  }
  return rules;
}

function applyInlineStyles(html, rules) {
  // For each rule, find matching elements and merge styles
  // Server-side approach: target common class/id/tag patterns
  let result = html;

  rules.forEach(({ selector, props }) => {
    if (!props || Object.keys(props).length === 0) return;
    const styleStr = Object.entries(props).map(([k, v]) => `${k}:${v}`).join(';');

    // Handle class selectors like .classname
    const classMatch = selector.match(/^\.([a-zA-Z0-9_-]+)$/);
    if (classMatch) {
      const className = classMatch[1];
      // Add/merge inline style to elements with this class
      result = result.replace(
        new RegExp(`(<[a-zA-Z][^>]*\\bclass="[^"]*\\b${escapeRe(className)}\\b[^"]*"[^>]*)(style="([^"]*)")?`, 'g'),
        (match, before, styleAttr, existingStyle) => {
          if (existingStyle) {
            return `${before}style="${mergeStyles(styleStr, existingStyle)}"`;
          }
          // Insert style attribute
          return before.replace(/(\s*\/?>)$/, ` style="${styleStr}"$1`);
        }
      );
    }

    // Handle element selectors like p, td, h1, etc.
    const tagMatch = selector.match(/^([a-zA-Z][a-zA-Z0-9]*)$/);
    if (tagMatch) {
      const tag = tagMatch[1].toLowerCase();
      result = result.replace(
        new RegExp(`(<${tag}(\\s[^>]*)?)(?:\\sstyle="([^"]*)")?(?=>|\\s)`, 'gi'),
        (match, open, rest, existingStyle) => {
          if (existingStyle !== undefined) {
            return `${open} style="${mergeStyles(styleStr, existingStyle)}"`;
          }
          return `${open} style="${styleStr}"`;
        }
      );
    }
  });

  return result;
}

function mergeStyles(base, override) {
  const parse = s => {
    const m = {};
    (s || '').split(';').forEach(p => {
      const i = p.indexOf(':');
      if (i > 0) {
        const k = p.slice(0, i).trim().toLowerCase();
        const v = p.slice(i + 1).trim();
        if (k && v) m[k] = v;
      }
    });
    return m;
  };
  return Object.entries(Object.assign({}, parse(base), parse(override)))
    .map(([k, v]) => `${k}:${v}`).join(';');
}

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function wrapForOutlook(bodyHtml, subject = '') {
  return `<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="x-apple-disable-message-reformatting">
<!--[if gte mso 9]><xml><o:OfficeDocumentSettings><o:AllowPNG/><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><![endif]-->
<!--[if mso]><style>table{border-collapse:collapse}td,th{border:none}p{margin:0}a{color:inherit}.ExternalClass{width:100%}</style><![endif]-->
<style>body,table,td,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}table,td{mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse}img{-ms-interpolation-mode:bicubic;border:0;height:auto}body{margin:0!important;padding:0!important;width:100%!important}a[x-apple-data-detectors]{color:inherit!important;text-decoration:none!important}</style>
</head>
<body style="margin:0;padding:0">
${bodyHtml}
</body>
</html>`;
}

// ─── API ROUTES ───────────────────────────────────────────────────────────────

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', from: FROM, smtp: SMTP.host + ':' + SMTP.port });
});

// Verify SMTP connection
app.get('/api/verify', async (req, res) => {
  try {
    const transporter = createTransport();
    await transporter.verify();
    res.json({ ok: true, message: 'SMTP connection verified successfully' });
  } catch (err) {
    res.status(500).json({ ok: false, message: err.message });
  }
});

// Send email
app.post('/api/send', upload.single('htmlFile'), async (req, res) => {
  try {
    const { to, cc, bcc, subject, htmlContent } = req.body;

    if (!to) return res.status(400).json({ ok: false, message: 'Recipient (To) is required' });
    if (!subject) return res.status(400).json({ ok: false, message: 'Subject is required' });

    // Get HTML — either from uploaded file or from textarea
    let rawHtml = '';
    if (req.file) {
      rawHtml = req.file.buffer.toString('utf-8');
    } else if (htmlContent) {
      rawHtml = htmlContent;
    } else {
      return res.status(400).json({ ok: false, message: 'No HTML content provided' });
    }

    // Inline CSS
    const inlined = inlineCss(rawHtml);

    // Extract body
    const bodyMatch = inlined.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    let bodyContent = bodyMatch ? bodyMatch[1] : inlined;

    const finalHtml = wrapForOutlook(bodyContent, subject);

    // Build text version (strip HTML tags)
    const textVersion = rawHtml
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 500) + '...';

    const mailOptions = {
      from: FROM,
      to: to.split(',').map(e => e.trim()).filter(Boolean).join(', '),
      subject: subject,
      html: finalHtml,
      text: textVersion,
    };

    if (cc) {
      mailOptions.cc = cc.split(',').map(e => e.trim()).filter(Boolean).join(', ');
    }
    if (bcc) {
      mailOptions.bcc = bcc.split(',').map(e => e.trim()).filter(Boolean).join(', ');
    }

    const transporter = createTransport();
    const info = await transporter.sendMail(mailOptions);

    res.json({
      ok: true,
      message: 'Email sent successfully',
      messageId: info.messageId,
      accepted: info.accepted,
      rejected: info.rejected
    });

  } catch (err) {
    console.error('Send error:', err);
    res.status(500).json({ ok: false, message: err.message });
  }
});

// Preview — returns processed HTML for iframe
app.post('/api/preview', upload.single('htmlFile'), (req, res) => {
  try {
    let rawHtml = '';
    if (req.file) {
      rawHtml = req.file.buffer.toString('utf-8');
    } else if (req.body.htmlContent) {
      rawHtml = req.body.htmlContent;
    } else {
      return res.status(400).json({ ok: false, message: 'No HTML provided' });
    }

    const inlined = inlineCss(rawHtml);
    const bodyMatch = inlined.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    const bodyContent = bodyMatch ? bodyMatch[1] : inlined;
    const finalHtml = wrapForOutlook(bodyContent);

    res.json({ ok: true, html: finalHtml });
  } catch (err) {
    res.status(500).json({ ok: false, message: err.message });
  }
});

// Serve UI
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '.', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  ANABION Mailer running at http://localhost:${PORT}`);
  console.log(`  SMTP: ${SMTP.host}:${SMTP.port}`);
  console.log(`  From: ${FROM}\n`);
});
