(function() {
  const translations = {
    zh: {
      page_title: "Excel文件比较工具",
      app_title: "Excel文件比较工具",
      tagline: "快速发现并对比两份Excel文档的差异",
      file1_title: "文件1（旧版本）",
      file2_title: "文件2（新版本）",
      drag_or_click: "拖拽文件到此处或点击上传",
      supported_formats: "支持 .xlsx, .xls 格式",
      compare_now: "直接比较",
      select_start_cell: "选择起始单元格",
      start_compare: "开始比较",
      differences: "差异条目",
      maximize_preview: "最大化预览",
      restore: "退出最大化",
      filter_all: "全部",
      filter_diff: "差异",
      filter_added: "新增",
      filter_deleted: "删除",
      filter_modified: "修改",
      filter_same: "相同",
      export_excel: "导出Excel",
      no_differences: "文件完全相同",
      no_differences_desc: "两份Excel文件没有发现任何差异",
      start_cell_label: "从单元格开始：",
      reset: "重置",
      loading: "正在比较文件，请稍候...",
      full_preview_title: "全文条目预览",
      fullscreen: "全屏",
      exit_fullscreen: "退出全屏",
      close: "关闭",
      swap_files: "交换文件位置",
      footer_text: "Excel文件比较工具 \u00a9 2024 | 所有处理均在浏览器本地完成，文件不会上传到服务器",
      file_compare_report: "Excel比较",
      alert_invalid_file: "请上传 Excel 文件 (.xlsx 或 .xls)",
      alert_parse_failed: "文件解析失败",
      alert_upload_first: "请先上传并比较文件",
      alert_file_lost: "文件信息丢失，请重新上传",
      alert_compare_first: "请先执行文件比较",
      render_failed: "渲染失败",
      stats_item: "统计项",
      quantity: "数量",
      added: "新增",
      deleted: "删除",
      modified: "修改",
      old_file: "旧文件",
      new_file: "新文件",
      generated_time: "生成时间",
      color_legend: "颜色说明",
      meaning: "含义",
      green_meaning: "新文件中新增的内容",
      red_meaning: "旧文件中有，新文件中已删除的内容",
      yellow_bold: "修改后的新值",
      yellow_italic: "修改前的旧值",
      old_added: "旧文件中的新增内容",
      new_added: "新文件中的新增内容",
      new_deleted: "新文件中的已删除内容",
      new_modified: "新文件中的修改内容",
      old_modified: "旧文件中的修改内容",
      max_preview_preview: "最大化预览",
      max_preview_exit: "退出最大化",
      no_data: "无数据",
      more_rows: "... 还有 {count} 行"
    },
    en: {
      page_title: "Excel Compare Tool",
      app_title: "Excel Compare Tool",
      tagline: "Quickly compare two Excel documents and highlight differences",
      file1_title: "File 1 (Old Version)",
      file2_title: "File 2 (New Version)",
      drag_or_click: "Drag files here or click to upload",
      supported_formats: "Supports .xlsx, .xls formats",
      compare_now: "Compare Now",
      select_start_cell: "Select Start Cell",
      start_compare: "Start Comparison",
      differences: "Differences",
      maximize_preview: "Maximize Preview",
      restore: "Restore",
      filter_all: "All",
      filter_diff: "Differences",
      filter_added: "Added",
      filter_deleted: "Deleted",
      filter_modified: "Modified",
      filter_same: "Unchanged",
      export_excel: "Export Excel",
      no_differences: "Files are Identical",
      no_differences_desc: "No differences found between the two Excel files",
      start_cell_label: "Start from cell:",
      reset: "Reset",
      loading: "Comparing files, please wait...",
      full_preview_title: "Full Preview",
      fullscreen: "Fullscreen",
      exit_fullscreen: "Exit Fullscreen",
      close: "Close",
      swap_files: "Swap Files",
      footer_text: "Excel Compare Tool \u00a9 2024 | All processing is done locally in your browser. Files are never uploaded to any server.",
      file_compare_report: "Excel Comparison",
      alert_invalid_file: "Please upload Excel files (.xlsx or .xls)",
      alert_parse_failed: "Failed to parse file",
      alert_upload_first: "Please upload files to compare first",
      alert_file_lost: "File information lost, please re-upload",
      alert_compare_first: "Please compare files first",
      render_failed: "Rendering failed",
      stats_item: "Stats",
      quantity: "Count",
      added: "Added",
      deleted: "Deleted",
      modified: "Modified",
      old_file: "Old File",
      new_file: "New File",
      generated_time: "Generated",
      color_legend: "Color Legend",
      meaning: "Meaning",
      green_meaning: "New content added in the new file",
      red_meaning: "Content deleted from the old file",
      yellow_bold: "Modified value (new)",
      yellow_italic: "Modified value (old)",
      old_added: "Added content in old file",
      new_added: "Added content in new file",
      new_deleted: "Deleted content in new file",
      new_modified: "Modified content in new file",
      old_modified: "Modified content in old file",
      max_preview_preview: "Maximize Preview",
      max_preview_exit: "Exit Maximize",
      no_data: "No data",
      more_rows: "... {count} more rows"
    },
    es: {
      page_title: "Herramienta de Comparaci\u00f3n de Excel",
      app_title: "Herramienta de Comparaci\u00f3n de Excel",
      tagline: "Compare r\u00e1pidamente dos documentos de Excel y resalte las diferencias",
      file1_title: "Archivo 1 (Versi\u00f3n anterior)",
      file2_title: "Archivo 2 (Versi\u00f3n nueva)",
      drag_or_click: "Arrastre archivos aqu\u00ed o haga clic para cargar",
      supported_formats: "Compatible con formatos .xlsx, .xls",
      compare_now: "Comparar ahora",
      select_start_cell: "Seleccionar celda inicial",
      start_compare: "Iniciar comparaci\u00f3n",
      differences: "Diferencias",
      maximize_preview: "Maximizar vista previa",
      restore: "Restaurar",
      filter_all: "Todos",
      filter_diff: "Diferencias",
      filter_added: "Agregado",
      filter_deleted: "Eliminado",
      filter_modified: "Modificado",
      filter_same: "Sin cambios",
      export_excel: "Exportar Excel",
      no_differences: "Los archivos son id\u00e9nticos",
      no_differences_desc: "No se encontraron diferencias entre los dos archivos de Excel",
      start_cell_label: "Empezar desde la celda:",
      reset: "Reiniciar",
      loading: "Comparando archivos, por favor espere...",
      full_preview_title: "Vista previa completa",
      fullscreen: "Pantalla completa",
      exit_fullscreen: "Salir de pantalla completa",
      close: "Cerrar",
      swap_files: "Intercambiar archivos",
      footer_text: "Herramienta de Comparaci\u00f3n de Excel \u00a9 2024 | Todo el procesamiento se realiza localmente en su navegador. Los archivos nunca se suben a ning\u00fan servidor.",
      file_compare_report: "Comparaci\u00f3n de Excel",
      alert_invalid_file: "Por favor, suba archivos de Excel (.xlsx o .xls)",
      alert_parse_failed: "Error al analizar el archivo",
      alert_upload_first: "Por favor, suba archivos para comparar primero",
      alert_file_lost: "Informaci\u00f3n del archivo perdida, vuelva a subir",
      alert_compare_first: "Por favor, compare archivos primero",
      render_failed: "Error al renderizar",
      stats_item: "Estad\u00edsticas",
      quantity: "Cantidad",
      added: "Agregado",
      deleted: "Eliminado",
      modified: "Modificado",
      old_file: "Archivo anterior",
      new_file: "Archivo nuevo",
      generated_time: "Generado",
      color_legend: "Leyenda de colores",
      meaning: "Significado",
      green_meaning: "Contenido nuevo agregado en el archivo nuevo",
      red_meaning: "Contenido eliminado del archivo anterior",
      yellow_bold: "Valor modificado (nuevo)",
      yellow_italic: "Valor modificado (anterior)",
      old_added: "Contenido agregado en archivo anterior",
      new_added: "Contenido agregado en archivo nuevo",
      new_deleted: "Contenido eliminado en archivo nuevo",
      new_modified: "Contenido modificado en archivo nuevo",
      old_modified: "Contenido modificado en archivo anterior",
      max_preview_preview: "Maximizar vista previa",
      max_preview_exit: "Salir de maximizar",
      no_data: "Sin datos",
      more_rows: "... {count} filas m\u00e1s"
    },
    de: {
      page_title: "Excel-Vergleichstool",
      app_title: "Excel-Vergleichstool",
      tagline: "Vergleichen Sie schnell zwei Excel-Dokumente und heben Sie Unterschiede hervor",
      file1_title: "Datei 1 (Alte Version)",
      file2_title: "Datei 2 (Neue Version)",
      drag_or_click: "Dateien hierher ziehen oder klicken zum Hochladen",
      supported_formats: "Unterst\u00fctzt .xlsx, .xls Formate",
      compare_now: "Jetzt vergleichen",
      select_start_cell: "Startzelle ausw\u00e4hlen",
      start_compare: "Vergleich starten",
      differences: "Unterschiede",
      maximize_preview: "Vorschau maximieren",
      restore: "Wiederherstellen",
      filter_all: "Alle",
      filter_diff: "Unterschiede",
      filter_added: "Hinzugef\u00fcgt",
      filter_deleted: "Gel\u00f6scht",
      filter_modified: "Ge\u00e4ndert",
      filter_same: "Unver\u00e4ndert",
      export_excel: "Excel exportieren",
      no_differences: "Dateien sind identisch",
      no_differences_desc: "Es wurden keine Unterschiede zwischen den beiden Excel-Dateien gefunden",
      start_cell_label: "Starten ab Zelle:",
      reset: "Zur\u00fccksetzen",
      loading: "Dateien werden verglichen, bitte warten...",
      full_preview_title: "Vollvorschau",
      fullscreen: "Vollbild",
      exit_fullscreen: "Vollbild beenden",
      close: "Schlie\u00dfen",
      swap_files: "Dateien tauschen",
      footer_text: "Excel-Vergleichstool \u00a9 2024 | Alle Verarbeitung erfolgt lokal in Ihrem Browser. Dateien werden niemals auf einen Server hochgeladen.",
      file_compare_report: "Excel-Vergleich",
      alert_invalid_file: "Bitte laden Sie Excel-Dateien hoch (.xlsx oder .xls)",
      alert_parse_failed: "Datei konnte nicht analysiert werden",
      alert_upload_first: "Bitte laden Sie zuerst Dateien zum Vergleichen hoch",
      alert_file_lost: "Dateiinformationen verloren, bitte erneut hochladen",
      alert_compare_first: "Bitte vergleichen Sie zuerst die Dateien",
      render_failed: "Rendern fehlgeschlagen",
      stats_item: "Statistiken",
      quantity: "Anzahl",
      added: "Hinzugef\u00fcgt",
      deleted: "Gel\u00f6scht",
      modified: "Ge\u00e4ndert",
      old_file: "Alte Datei",
      new_file: "Neue Datei",
      generated_time: "Generiert",
      color_legend: "Farb-Legende",
      meaning: "Bedeutung",
      green_meaning: "Neuer Inhalt in der neuen Datei hinzugef\u00fcgt",
      red_meaning: "Inhalt aus der alten Datei gel\u00f6scht",
      yellow_bold: "Ge\u00e4nderter Wert (neu)",
      yellow_italic: "Ge\u00e4nderter Wert (alt)",
      old_added: "Hinzugef\u00fcgter Inhalt in alter Datei",
      new_added: "Hinzugef\u00fcgter Inhalt in neuer Datei",
      new_deleted: "Gel\u00f6schter Inhalt in neuer Datei",
      new_modified: "Ge\u00e4nderter Inhalt in neuer Datei",
      old_modified: "Ge\u00e4nderter Inhalt in alter Datei",
      max_preview_preview: "Vorschau maximieren",
      max_preview_exit: "Maximieren beenden",
      no_data: "Keine Daten",
      more_rows: "... {count} weitere Zeilen"
    },
    fr: {
      page_title: "Outil de Comparaison Excel",
      app_title: "Outil de Comparaison Excel",
      tagline: "Comparez rapidement deux documents Excel et mettez en surbrillance les diff\u00e9rences",
      file1_title: "Fichier 1 (Ancienne version)",
      file2_title: "Fichier 2 (Nouvelle version)",
      drag_or_click: "Glissez des fichiers ici ou cliquez pour t\u00e9l\u00e9charger",
      supported_formats: "Prend en charge les formats .xlsx, .xls",
      compare_now: "Comparer maintenant",
      select_start_cell: "S\u00e9lectionner la cellule de d\u00e9part",
      start_compare: "Commencer la comparaison",
      differences: "Diff\u00e9rences",
      maximize_preview: "Maximiser l'aper\u00e7u",
      restore: "Restaurer",
      filter_all: "Tous",
      filter_diff: "Diff\u00e9rences",
      filter_added: "Ajout\u00e9",
      filter_deleted: "Supprim\u00e9",
      filter_modified: "Modifi\u00e9",
      filter_same: "Identique",
      export_excel: "Exporter Excel",
      no_differences: "Les fichiers sont identiques",
      no_differences_desc: "Aucune diff\u00e9rence trouv\u00e9e entre les deux fichiers Excel",
      start_cell_label: "Commencer \u00e0 partir de la cellule:",
      reset: "R\u00e9initialiser",
      loading: "Comparaison des fichiers, veuillez patienter...",
      full_preview_title: "Aper\u00e7u complet",
      fullscreen: "Plein \u00e9cran",
      exit_fullscreen: "Quitter le plein \u00e9cran",
      close: "Fermer",
      swap_files: "\u00c9changer les fichiers",
      footer_text: "Outil de Comparaison Excel \u00a9 2024 | Tout le traitement est effectu\u00e9 localement dans votre navigateur. Les fichiers ne sont jamais t\u00e9l\u00e9charg\u00e9s sur un serveur.",
      file_compare_report: "Comparaison Excel",
      alert_invalid_file: "Veuillez t\u00e9l\u00e9charger des fichiers Excel (.xlsx ou .xls)",
      alert_parse_failed: "\u00c9chec de l'analyse du fichier",
      alert_upload_first: "Veuillez d'abord t\u00e9l\u00e9charger des fichiers \u00e0 comparer",
      alert_file_lost: "Informations du fichier perdues, veuillez re-t\u00e9l\u00e9charger",
      alert_compare_first: "Veuillez d'abord comparer les fichiers",
      render_failed: "\u00c9chec du rendu",
      stats_item: "Statistiques",
      quantity: "Quantit\u00e9",
      added: "Ajout\u00e9",
      deleted: "Supprim\u00e9",
      modified: "Modifi\u00e9",
      old_file: "Ancien fichier",
      new_file: "Nouveau fichier",
      generated_time: "G\u00e9n\u00e9r\u00e9",
      color_legend: "L\u00e9gende des couleurs",
      meaning: "Signification",
      green_meaning: "Nouveau contenu ajout\u00e9 dans le nouveau fichier",
      red_meaning: "Contenu supprim\u00e9 de l'ancien fichier",
      yellow_bold: "Valeur modifi\u00e9e (nouvelle)",
      yellow_italic: "Valeur modifi\u00e9e (ancienne)",
      old_added: "Contenu ajout\u00e9 dans l'ancien fichier",
      new_added: "Contenu ajout\u00e9 dans le nouveau fichier",
      new_deleted: "Contenu supprim\u00e9 dans le nouveau fichier",
      new_modified: "Contenu modifi\u00e9 dans le nouveau fichier",
      old_modified: "Contenu modifi\u00e9 dans l'ancien fichier",
      max_preview_preview: "Maximiser l'aper\u00e7u",
      max_preview_exit: "Quitter l'agrandissement",
      no_data: "Aucune donn\u00e9e",
      more_rows: "... {count} lignes suppl\u00e9mentaires"
    }
  };

  let currentLang = localStorage.getItem('excel-compare-lang');
  if (!currentLang) {
    const browserLang = navigator.language.slice(0, 2);
    currentLang = ['en', 'zh', 'es', 'de', 'fr'].includes(browserLang) ? browserLang : 'en';
    localStorage.setItem('excel-compare-lang', currentLang);
  }

  function applyLanguage() {
    const t = translations[currentLang];
    if (!t) return;

    document.querySelectorAll('[data-i18n]').forEach(function(el) {
      const key = el.getAttribute('data-i18n');
      if (t[key] !== undefined) {
        el.textContent = t[key];
      }
    });

    document.querySelectorAll('[data-i18n-placeholder]').forEach(function(el) {
      const key = el.getAttribute('data-i18n-placeholder');
      if (t[key] !== undefined) {
        el.setAttribute('placeholder', t[key]);
      }
    });

    document.querySelectorAll('[data-i18n-title]').forEach(function(el) {
      const key = el.getAttribute('data-i18n-title');
      if (t[key] !== undefined) {
        el.setAttribute('title', t[key]);
      }
    });

    document.querySelectorAll('[data-i18n-html]').forEach(function(el) {
      const key = el.getAttribute('data-i18n-html');
      if (t[key] !== undefined) {
        el.innerHTML = t[key];
      }
    });

    document.querySelector('meta[name="description"]')?.setAttribute('content', t.tagline || '');
    document.documentElement.lang = currentLang === 'zh' ? 'zh-CN' : currentLang;

    var langSelector = document.getElementById('langSelector');
    if (langSelector) langSelector.value = currentLang;
  }

  function setLanguage(lang) {
    if (!translations[lang]) return;
    currentLang = lang;
    localStorage.setItem('excel-compare-lang', lang);
    applyLanguage();
  }

  window.i18n = {
    t: function(key, params) {
      var t = translations[currentLang] || translations['en'];
      var text = t[key] || key;
      if (params) {
        for (var k in params) {
          if (params.hasOwnProperty(k)) {
            text = text.replace('{' + k + '}', params[k]);
          }
        }
      }
      return text;
    },
    setLanguage: setLanguage,
    getCurrentLang: function() { return currentLang; },
    applyLanguage: applyLanguage
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', applyLanguage);
  } else {
    applyLanguage();
  }
})();
