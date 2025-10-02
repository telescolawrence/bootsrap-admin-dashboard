(function() {
  'use strict';

  const exportButtons = document.querySelectorAll('[data-export-target]');

  if (!exportButtons.length) {
    return;
  }

  const sanitiseTableForExport = (table) => {
    if (!table) {
      return '';
    }

    const clone = table.cloneNode(true);

    clone.querySelectorAll('[data-export-exclude]')
      .forEach((element) => element.remove());

    clone.querySelectorAll('button, .btn, .badge, .sr-only').forEach((element) => {
      if (element.closest('th') || element.closest('td')) {
        element.replaceWith(element.textContent.trim());
      } else {
        element.remove();
      }
    });

    return clone.outerHTML;
  };

  const buildWorkbookHtml = (options) => {
    const { mainTable, balanceTable, title } = options;
    const sections = [];

    if (title) {
      const columnCount = mainTable && mainTable.tHead
        ? mainTable.tHead.rows[mainTable.tHead.rows.length - 1].cells.length
        : (mainTable && mainTable.rows[0] ? mainTable.rows[0].cells.length : 1);
      sections.push(
        `<table><thead><tr><th colspan="${columnCount}" style="text-align:left;font-weight:bold;">${title}</th></tr></thead></table>`
      );
    }

    sections.push(sanitiseTableForExport(mainTable));

    if (balanceTable) {
      sections.push('<table><tbody><tr><td style="height:12px"></td></tr></tbody></table>');
      sections.push(sanitiseTableForExport(balanceTable));
    }

    return `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
        <head>
          <meta charset="UTF-8">
        </head>
        <body>
          ${sections.join('\n')}
        </body>
      </html>
    `;
  };

  const triggerDownload = (html, filename) => {
    const blob = new Blob(['\ufeff' + html], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename || 'supplier-summary.xls';
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
  };

  exportButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const mainSelector = button.getAttribute('data-export-target');
      if (!mainSelector) {
        return;
      }

      const mainTable = document.querySelector(mainSelector);
      if (!mainTable) {
        console.warn('Supplier summary export: Unable to find summary table using selector', mainSelector);
        return;
      }

      const balanceSelector = button.getAttribute('data-balance-table');
      const balanceTable = balanceSelector ? document.querySelector(balanceSelector) : null;
      const filename = button.getAttribute('data-export-filename') || 'supplier-summary.xls';
      const title = button.getAttribute('data-export-title') || '';

      const workbookHtml = buildWorkbookHtml({ mainTable, balanceTable, title });
      triggerDownload(workbookHtml, filename);
    });
  });
})();
