Object.assign(MakeItRed, {
    addExportPanelToWindow(window, toolbar, popupSet) {
        let doc = window.document;

        let exportButton = doc.createXULElement('toolbarbutton');
        exportButton.id = 'make-it-red-export-toolbar-button';
        exportButton.setAttribute('label', 'Export PDFs');
        exportButton.setAttribute('tooltiptext', 'Open PDF Export panel');
        exportButton.setAttribute('class', 'toolbarbutton-1');
        exportButton.setAttribute('image', this.rootURI + 'icon.png');
        toolbar.appendChild(exportButton);
        this.storeAddedElement(exportButton);

        let exportPanel = doc.createXULElement('panel');
        exportPanel.id = 'make-it-red-export-panel';
        exportPanel.setAttribute('type', 'arrow');
        exportPanel.setAttribute('orient', 'vertical');
        exportPanel.setAttribute('style', 'padding: 15px; width: 800px; max-height: 85vh; display: flex; flex-direction: column; box-sizing: border-box;');

        let exVbox = doc.createXULElement('vbox');
        exVbox.setAttribute('style', 'gap: 10px; flex: 1; display: flex; flex-direction: column; overflow: hidden;');
        exVbox.setAttribute('flex', '1');

        let exPathLabel = doc.createXULElement('label');
        exPathLabel.setAttribute('value', 'Export Folder:');

        let exPathRow = doc.createXULElement('hbox');
        exPathRow.setAttribute('style', 'gap: 8px; align-items: center;');

        let exPathInput = doc.createElementNS('http://www.w3.org/1999/xhtml', 'input');
        exPathInput.id = 'make-it-red-export-path';
        exPathInput.setAttribute('type', 'text');
        exPathInput.setAttribute('placeholder', 'Select export folder...');
        exPathInput.setAttribute('style', 'flex: 1; box-sizing: border-box; padding: 4px;');
        exPathInput.value = this.lastExportPath;

        exPathInput.addEventListener('change', () => {
            this.lastExportPath = exPathInput.value.trim();
        });

        let exBrowseBtn = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
        exBrowseBtn.className = 'make-it-red-html-btn';
        exBrowseBtn.textContent = 'Browse...';

        exBrowseBtn.addEventListener('click', async () => {
            try {
                let pickedPath = await this.pickFolderPath(window, this.lastExportPath || exPathInput.value);
                if (pickedPath) {
                    exPathInput.value = pickedPath;
                    this.lastExportPath = pickedPath;
                }
            }
            catch (e) {
                this.log('File picker error: ' + e);
                exStatus.setAttribute('value', 'Error: Unable to open folder picker.');
            }
        });

        exPathRow.appendChild(exPathInput);
        exPathRow.appendChild(exBrowseBtn);

        let exActions = doc.createXULElement('hbox');
        exActions.setAttribute('style', 'gap: 8px;');

        let selectAllBtn = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
        selectAllBtn.className = 'make-it-red-html-btn';
        selectAllBtn.textContent = 'Select All';

        let deselectAllBtn = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
        deselectAllBtn.className = 'make-it-red-html-btn';
        deselectAllBtn.textContent = 'Deselect All';

        let doExportBtn = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
        doExportBtn.className = 'make-it-red-html-btn';
        doExportBtn.textContent = 'Export Selected PDFs';

        exActions.appendChild(selectAllBtn);
        exActions.appendChild(deselectAllBtn);
        exActions.appendChild(doExportBtn);

        let exStatus = doc.createXULElement('label');
        exStatus.setAttribute('value', 'Ready.');

        let exOutput = doc.createElementNS('http://www.w3.org/1999/xhtml', 'div');
        exOutput.id = 'make-it-red-export-output';
        exOutput.setAttribute('style', 'width: 100%; flex: 1; min-height: 300px; overflow-y: auto; background-color: #f9f9f9; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;');

        selectAllBtn.addEventListener('click', () => {
            let checkboxes = exOutput.querySelectorAll('.make-it-red-export-checkbox:not(:disabled)');
            checkboxes.forEach(cb => cb.checked = true);
        });

        deselectAllBtn.addEventListener('click', () => {
            let checkboxes = exOutput.querySelectorAll('.make-it-red-export-checkbox');
            checkboxes.forEach(cb => cb.checked = false);
        });

        doExportBtn.addEventListener('click', async () => {
            let targetPath = exPathInput.value.trim();
            if (!targetPath) {
                exStatus.setAttribute('value', 'Error: Please select an export folder first.');
                return;
            }

            let componentsObj = this.getComponents(window);
            if (!componentsObj?.classes || !componentsObj?.interfaces) {
                exStatus.setAttribute('value', 'Error: File APIs are unavailable in this context.');
                return;
            }
            let Ci = componentsObj.interfaces;

            let destDirFile;
            try {
                destDirFile = componentsObj.classes['@mozilla.org/file/local;1'].createInstance(Ci.nsIFile);
                destDirFile.initWithPath(targetPath);
                if (!destDirFile.exists() || !destDirFile.isDirectory()) {
                    exStatus.setAttribute('value', 'Error: Target directory does not exist.');
                    return;
                }
            }
            catch (e) {
                exStatus.setAttribute('value', 'Error: Invalid directory path.');
                return;
            }

            let checkboxes = exOutput.querySelectorAll('.make-it-red-export-checkbox:checked');
            if (checkboxes.length === 0) {
                exStatus.setAttribute('value', 'No items selected for export.');
                return;
            }

            doExportBtn.disabled = true;
            exStatus.setAttribute('value', 'Exporting...');
            let successCount = 0;

            for (let cb of checkboxes) {
                let itemID = parseInt(cb.value, 10);
                let item = Zotero.Items.get(itemID);
                if (!item) continue;

                let attIDs = item.getAttachments();
                let pdfPath = null;
                for (let id of attIDs) {
                    let att = Zotero.Items.get(id);
                    if (att && att.attachmentContentType === 'application/pdf') {
                        pdfPath = await att.getFilePathAsync();
                        break;
                    }
                }

                if (pdfPath) {
                    try {
                        let srcFile = componentsObj.classes['@mozilla.org/file/local;1'].createInstance(Ci.nsIFile);
                        srcFile.initWithPath(pdfPath);
                        if (srcFile.exists()) {
                            srcFile.copyToUnique(destDirFile, srcFile.leafName);
                            successCount++;
                        }
                    }
                    catch (e) {
                        this.log(`Failed to copy PDF for item ${itemID}: ${e}`);
                    }
                }
            }

            doExportBtn.disabled = false;
            exStatus.setAttribute('value', `Done: Exported ${successCount} PDFs to folder.`);
        });

        exportButton.addEventListener('command', async () => {
            if (exportPanel.state === 'open' || exportPanel.state === 'showing') {
                exportPanel.hidePopup();
                return;
            }

            exPathInput.value = this.lastExportPath;
            exportPanel.openPopup(exportButton, 'after_start', 0, 6, false, false);
            exStatus.setAttribute('value', 'Loading items...');
            try {
                let count = await this.renderExportItems(window, exOutput);
                exStatus.setAttribute('value', `Ready. Loaded ${count} items.`);
            }
            catch (e) {
                this.log(`Failed to render export items: ${e}`);
                exOutput.innerHTML = `<div style="padding: 12px; color: #d32f2f;">Failed to load items: ${this.escapeHTML(e.message || String(e))}</div>`;
                exStatus.setAttribute('value', 'Failed to load items.');
            }
        });

        exVbox.appendChild(exPathLabel);
        exVbox.appendChild(exPathRow);
        exVbox.appendChild(exActions);
        exVbox.appendChild(exStatus);
        exVbox.appendChild(exOutput);
        exportPanel.appendChild(exVbox);

        popupSet.appendChild(exportPanel);
        this.storeAddedElement(exportPanel);
    },

    async renderExportItems(window, container) {
        try {
            let items = [];
            try {
                items = this.getCurrentVisibleRegularItems(window);
            }
            catch (e) {
                this.log(`getCurrentVisibleRegularItems threw: ${e}`);
                items = [];
            }

            if (items.length === 0) {
                try {
                    let fallback = window?.ZoteroPane?.getSelectedItems?.() || [];
                    items = fallback.filter(i => i && typeof i.isRegularItem === 'function' && i.isRegularItem());
                }
                catch (e) {
                    this.log(`getSelectedItems fallback failed: ${e}`);
                }
            }

            if (items.length === 0) {
                container.innerHTML = `
					<div style="padding: 30px; text-align: center; color: #555;">
						<h3 style="margin-bottom: 10px;">目前未找到待导出的条目</h3>
						<p>请先在中间列表显示出条目（例如选中某个分类或执行检索），</p>
						<p>然后再次打开导出面板。</p>
					</div>
				`;
                return 0;
            }

            let html = `
			<table style="width: 100%; border-collapse: collapse; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 13px;">
				<thead style="position: sticky; top: 0; background-color: #eaeaea; z-index: 1; box-shadow: 0 1px 2px rgba(0,0,0,0.1);">
					<tr>
						<th style="padding: 8px; border: 1px solid #ccc; width: 5%; text-align: center;"></th>
						<th style="padding: 8px; border: 1px solid #ccc; width: 50%; text-align: left;">Title</th>
						<th style="padding: 8px; border: 1px solid #ccc; width: 30%; text-align: left;">Creator</th>
						<th style="padding: 8px; border: 1px solid #ccc; width: 15%; text-align: center;">Has PDF</th>
					</tr>
				</thead>
				<tbody>
			`;

            for (let item of items) {
                try {
                    if (!item || typeof item.getAttachments !== 'function') continue;

                    let hasPdf = false;
                    let attIDs = item.getAttachments();
                    for (let id of attIDs) {
                        let att = Zotero.Items.get(id);
                        if (att && att.attachmentContentType === 'application/pdf') {
                            hasPdf = true;
                            break;
                        }
                    }

                    let title = this.escapeHTML(item.getField('title') || 'Untitled');
                    let creator = this.escapeHTML(item.firstCreator || '');
                    let pdfStatus = hasPdf ? '<span style="color: #2e7d32; font-weight: bold;">Yes</span>' : '<span style="color: #d32f2f;">No</span>';
                    let disabled = hasPdf ? '' : 'disabled';
                    let checked = hasPdf ? 'checked' : '';

                    html += `
						<tr style="border-bottom: 1px solid #eee; background-color: #fff;">
							<td style="padding: 8px; text-align: center; border: 1px solid #eee;">
								<input type="checkbox" class="make-it-red-export-checkbox" value="${item.id}" ${disabled} ${checked}>
							</td>
							<td style="padding: 8px; color: #333; border: 1px solid #eee;">${title}</td>
							<td style="padding: 8px; color: #666; border: 1px solid #eee;">${creator}</td>
							<td style="padding: 8px; text-align: center; border: 1px solid #eee;">${pdfStatus}</td>
						</tr>
					`;
                }
                catch (e) {
                    this.log(`Skipping malformed item in export list: ${e}`);
                }
            }

            html += `</tbody></table>`;
            container.innerHTML = html;
            return items.length;
        }
        catch (e) {
            let details = String(e?.message || e || 'Unknown error');
            this.log(`renderExportItems failed: ${details} :: ${e?.stack || ''}`);
            container.innerHTML = `<div style="padding: 12px; color: #d32f2f;">Failed to load items: ${this.escapeHTML(details)}</div>`;
            return 0;
        }
    },

    getCurrentVisibleRegularItems(window) {
        let pane = window?.ZoteroPane;
        if (!pane) return [];

        let items = [];
        let addRegularItems = (candidate) => {
            if (!candidate) return;
            for (let item of candidate) {
                if (item && typeof item.isRegularItem === 'function' && item.isRegularItem()) {
                    items.push(item);
                }
            }
        };

        let addItemsFromUnknownList = (candidate) => {
            if (!candidate || !candidate.length) return;

            if (candidate[0] && typeof candidate[0].isRegularItem === 'function') {
                addRegularItems(candidate);
                return;
            }

            for (let raw of candidate) {
                let id = null;
                if (typeof raw === 'number' && Number.isInteger(raw)) {
                    id = raw;
                }
                else if (typeof raw === 'string') {
                    let parsed = Number.parseInt(raw, 10);
                    if (Number.isInteger(parsed)) id = parsed;
                }
                else if (raw && typeof raw === 'object' && Number.isInteger(raw.id)) {
                    id = raw.id;
                }

                if (id == null) continue;

                try {
                    let item = Zotero.Items.get(id);
                    if (item && typeof item.isRegularItem === 'function' && item.isRegularItem()) {
                        items.push(item);
                    }
                }
                catch (e) {
                    this.log(`Zotero.Items.get(${id}) failed: ${e}`);
                }
            }
        };

        try {
            if (typeof pane.getSortedItems === 'function') {
                let sortedItems = pane.getSortedItems() || [];
                addItemsFromUnknownList(sortedItems);
            }
        }
        catch (e) {
            this.log(`getSortedItems() failed: ${e}`);
        }

        if (items.length === 0) {
            try {
                if (typeof pane.getSortedItems === 'function') {
                    let sortedIDs = pane.getSortedItems(true) || [];
                    addItemsFromUnknownList(sortedIDs);
                }
            }
            catch (e) {
                this.log(`getSortedItems(true) failed: ${e}`);
            }
        }

        if (items.length === 0) {
            try {
                let collection = pane.getSelectedCollection?.();
                if (collection && typeof collection.getChildItems === 'function') {
                    addRegularItems(collection.getChildItems());
                }
            }
            catch (e) {
                this.log(`getSelectedCollection/getChildItems failed: ${e}`);
            }
        }

        if (items.length === 0) {
            try {
                if (typeof pane.getSelectedItems === 'function') {
                    addRegularItems(pane.getSelectedItems());
                }
            }
            catch (e) {
                this.log(`getSelectedItems failed: ${e}`);
            }
        }

        let dedup = new Map();
        for (let item of items) {
            if (item?.id != null) dedup.set(item.id, item);
        }
        return Array.from(dedup.values());
    },

    async pickFolderPath(window, defaultPath = '') {
        let componentsObj = this.getComponents(window);

        if (componentsObj?.classes && componentsObj?.interfaces) {
            let Ci = componentsObj.interfaces;
            let fp = componentsObj.classes['@mozilla.org/filepicker;1'].createInstance(Ci.nsIFilePicker);
            fp.init(window, 'Select Export Directory', Ci.nsIFilePicker.modeGetFolder);

            if (defaultPath) {
                try {
                    let dir = componentsObj.classes['@mozilla.org/file/local;1'].createInstance(Ci.nsIFile);
                    dir.initWithPath(defaultPath);
                    if (dir.exists() && dir.isDirectory()) {
                        fp.displayDirectory = dir;
                    }
                }
                catch (e) {
                    this.log(`Set default picker directory failed: ${e}`);
                }
            }

            return await new Promise((resolve) => {
                fp.open((rv) => {
                    if (rv === Ci.nsIFilePicker.returnOK && fp.file) {
                        resolve(fp.file.path);
                        return;
                    }
                    resolve('');
                });
            });
        }

        try {
            if (typeof ChromeUtils !== 'undefined' && typeof ChromeUtils.importESModule === 'function') {
                let { FilePicker } = ChromeUtils.importESModule('resource://gre/modules/FilePicker.sys.mjs');
                let picker = new FilePicker(window.browsingContext);
                picker.init(window, 'Select Export Directory', picker.modeGetFolder);
                let rv = await picker.show();
                if (rv === picker.returnOK && picker.file) {
                    return picker.file.path;
                }
            }
        }
        catch (e) {
            this.log(`FilePicker.sys.mjs fallback failed: ${e}`);
        }

        try {
            if (typeof Services !== 'undefined' && Services.prompt) {
                let value = { value: String(defaultPath || '') };
                let ok = Services.prompt.prompt(window, 'Export Folder', 'Enter export folder path:', value, null, {});
                if (ok) {
                    return String(value.value || '').trim();
                }
            }
        }
        catch (e) {
            this.log(`Services.prompt fallback failed: ${e}`);
        }

        try {
            let manualPath = window.prompt('Folder picker is unavailable in this context. Please enter export folder path:', defaultPath || '');
            return String(manualPath || '').trim();
        }
        catch (e) {
            this.log(`window.prompt fallback failed: ${e}`);
            return '';
        }
    },

    getComponents(window) {
        return (typeof Components !== 'undefined' && Components)
            || window?.Components
            || null;
    },
});
