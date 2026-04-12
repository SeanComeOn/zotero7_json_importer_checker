Object.assign(MakeItRed, {
    addExportPanelToWindow(window, toolbar, popupSet) {
        let doc = window.document;

        let exportButton = doc.createXULElement('toolbarbutton');
        exportButton.id = 'wo-yao-yan-pai-export-toolbar-button';
        exportButton.setAttribute('label', 'Export PDFs');
        exportButton.setAttribute('tooltiptext', 'Open PDF Export panel');
        exportButton.setAttribute('class', 'toolbarbutton-1');
        exportButton.setAttribute('image', this.rootURI + 'icons/icon.png');
        toolbar.appendChild(exportButton);
        this.storeAddedElement(exportButton);

        let exportPanel = doc.createXULElement('panel');
        exportPanel.id = 'wo-yao-yan-pai-export-panel';
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
        exPathInput.id = 'wo-yao-yan-pai-export-path';
        exPathInput.setAttribute('type', 'text');
        exPathInput.setAttribute('placeholder', 'Select export folder...');
        exPathInput.setAttribute('style', 'flex: 1; box-sizing: border-box; padding: 4px;');
        exPathInput.value = this.lastExportPath;

        exPathInput.addEventListener('change', () => {
            this.lastExportPath = exPathInput.value.trim();
        });

        let exBrowseBtn = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
        exBrowseBtn.className = 'wo-yao-yan-pai-html-btn';
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
        selectAllBtn.className = 'wo-yao-yan-pai-html-btn';
        selectAllBtn.textContent = 'Select All';

        let deselectAllBtn = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
        deselectAllBtn.className = 'wo-yao-yan-pai-html-btn';
        deselectAllBtn.textContent = 'Deselect All';

        let doExportBtn = doc.createElementNS('http://www.w3.org/1999/xhtml', 'button');
        doExportBtn.className = 'wo-yao-yan-pai-html-btn';
        doExportBtn.textContent = 'Export Selected PDFs';

        exActions.appendChild(selectAllBtn);
        exActions.appendChild(deselectAllBtn);
        exActions.appendChild(doExportBtn);

        let exStatus = doc.createXULElement('label');
        exStatus.setAttribute('value', 'Ready.');

        let exOutput = doc.createElementNS('http://www.w3.org/1999/xhtml', 'div');
        exOutput.id = 'wo-yao-yan-pai-export-output';
        exOutput.setAttribute('style', 'width: 100%; flex: 1; min-height: 300px; overflow-y: auto; background-color: #f9f9f9; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;');

        selectAllBtn.addEventListener('click', () => {
            let checkboxes = exOutput.querySelectorAll('.wo-yao-yan-pai-export-checkbox:not(:disabled)');
            checkboxes.forEach(cb => cb.checked = true);
        });

        deselectAllBtn.addEventListener('click', () => {
            let checkboxes = exOutput.querySelectorAll('.wo-yao-yan-pai-export-checkbox');
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

            let checkboxes = exOutput.querySelectorAll('.wo-yao-yan-pai-export-checkbox:checked');
            if (checkboxes.length === 0) {
                exStatus.setAttribute('value', 'No items selected for export.');
                return;
            }

            doExportBtn.disabled = true;
            exStatus.setAttribute('value', 'Exporting...');
            let successCount = 0;
            let skippedNoPathCount = 0;
            let skippedMissingFileCount = 0;
            let copyFailedCount = 0;
            let firstCopyError = '';

            for (let cb of checkboxes) {
                let attachmentID = parseInt(cb.value, 10);
                let attachment = Zotero.Items.get(attachmentID);
                if (!attachment || !attachment.isAttachment?.()) continue;

                try {
                    let srcFile = await this.getAttachmentSourceFile(attachment, window);
                    if (!srcFile) {
                        skippedNoPathCount++;
                        continue;
                    }

                    if (typeof srcFile.exists === 'function' && srcFile.exists()) {
                        await this.copyFileToDirectoryUnique(srcFile, destDirFile, window);
                        successCount++;
                    }
                    else {
                        skippedMissingFileCount++;
                    }
                }
                catch (e) {
                    copyFailedCount++;
                    let dbgPath = await this.getAttachmentFilePath(attachment);
                    if (!firstCopyError) {
                        firstCopyError = String(e?.message || e || 'copy failed');
                    }
                    this.log(`Failed to copy PDF attachment ${attachmentID} (path=${dbgPath || 'N/A'}): ${e}`);
                }
            }

            doExportBtn.disabled = false;
            let selectedCount = checkboxes.length;
            let skippedCount = skippedNoPathCount + skippedMissingFileCount;
            if (successCount === 0) {
                exStatus.setAttribute(
                    'value',
                    `Done: Exported 0/${selectedCount}. Skipped ${skippedCount} (no local path: ${skippedNoPathCount}, missing file: ${skippedMissingFileCount}), failed ${copyFailedCount}.`
                );
            }
            else {
                exStatus.setAttribute(
                    'value',
                    `Done: Exported ${successCount}/${selectedCount}. Skipped ${skippedCount} (no local path: ${skippedNoPathCount}, missing file: ${skippedMissingFileCount}), failed ${copyFailedCount}.`
                );
            }

            if (firstCopyError) {
                exStatus.setAttribute(
                    'value',
                    `${exStatus.getAttribute('value')} First error: ${firstCopyError}`
                );
            }
        });

        exportButton.addEventListener('command', async () => {
            if (exportPanel.state === 'open' || exportPanel.state === 'showing') {
                exportPanel.hidePopup();
                return;
            }

            exPathInput.value = this.lastExportPath;
            exportPanel.openPopup(exportButton, 'after_start', 0, 6, false, false);
            exStatus.setAttribute('value', 'Loading PDFs from current view...');
            try {
                let count = await this.renderExportItems(window, exOutput);
                exStatus.setAttribute('value', `Ready. Loaded ${count} PDFs.`);
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
            let doc = container.ownerDocument || window?.document;
            let pdfEntries = await this.getCurrentVisiblePDFEntries(window);

            if (pdfEntries.length === 0) {
                container.textContent = '';
                let msgWrap = doc.createElementNS('http://www.w3.org/1999/xhtml', 'div');
                msgWrap.setAttribute('style', 'padding: 30px; text-align: center; color: #555;');

                let h3 = doc.createElementNS('http://www.w3.org/1999/xhtml', 'h3');
                h3.setAttribute('style', 'margin-bottom: 10px;');
                h3.textContent = '目前未找到可导出的 PDF';

                let p1 = doc.createElementNS('http://www.w3.org/1999/xhtml', 'p');
                p1.textContent = '请先在中间列表显示出条目（例如选中某个分类或执行检索），';

                let p2 = doc.createElementNS('http://www.w3.org/1999/xhtml', 'p');
                p2.textContent = '然后再次打开导出面板。';

                msgWrap.appendChild(h3);
                msgWrap.appendChild(p1);
                msgWrap.appendChild(p2);
                container.appendChild(msgWrap);
                return 0;
            }

            container.textContent = '';

            let table = doc.createElementNS('http://www.w3.org/1999/xhtml', 'table');
            table.setAttribute('style', "width: 100%; border-collapse: collapse; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 13px;");

            let thead = doc.createElementNS('http://www.w3.org/1999/xhtml', 'thead');
            thead.setAttribute('style', 'position: sticky; top: 0; background-color: #eaeaea; z-index: 1; box-shadow: 0 1px 2px rgba(0,0,0,0.1);');
            let headRow = doc.createElementNS('http://www.w3.org/1999/xhtml', 'tr');

            let makeTH = (text, style) => {
                let th = doc.createElementNS('http://www.w3.org/1999/xhtml', 'th');
                th.setAttribute('style', style);
                th.textContent = text;
                return th;
            };

            headRow.appendChild(makeTH('', 'padding: 8px; border: 1px solid #ccc; width: 5%; text-align: center;'));
            headRow.appendChild(makeTH('PDF File', 'padding: 8px; border: 1px solid #ccc; width: 42%; text-align: left;'));
            headRow.appendChild(makeTH('Parent Title', 'padding: 8px; border: 1px solid #ccc; width: 38%; text-align: left;'));
            headRow.appendChild(makeTH('Creator', 'padding: 8px; border: 1px solid #ccc; width: 15%; text-align: left;'));

            thead.appendChild(headRow);
            table.appendChild(thead);

            let tbody = doc.createElementNS('http://www.w3.org/1999/xhtml', 'tbody');

            for (let entry of pdfEntries) {
                try {
                    let row = doc.createElementNS('http://www.w3.org/1999/xhtml', 'tr');
                    row.setAttribute('style', 'border-bottom: 1px solid #eee; background-color: #fff;');

                    let checkboxCell = doc.createElementNS('http://www.w3.org/1999/xhtml', 'td');
                    checkboxCell.setAttribute('style', 'padding: 8px; text-align: center; border: 1px solid #eee;');

                    let checkbox = doc.createElementNS('http://www.w3.org/1999/xhtml', 'input');
                    checkbox.setAttribute('type', 'checkbox');
                    checkbox.setAttribute('class', 'wo-yao-yan-pai-export-checkbox');
                    checkbox.setAttribute('value', String(entry.attachmentID));
                    checkbox.checked = true;
                    checkboxCell.appendChild(checkbox);
                    row.appendChild(checkboxCell);

                    let makeTD = (text, style) => {
                        let td = doc.createElementNS('http://www.w3.org/1999/xhtml', 'td');
                        td.setAttribute('style', style);
                        td.textContent = text;
                        return td;
                    };

                    row.appendChild(makeTD(String(entry.fileName || 'Untitled.pdf'), 'padding: 8px; color: #333; border: 1px solid #eee;'));
                    row.appendChild(makeTD(String(entry.parentTitle || '(No parent item)'), 'padding: 8px; color: #333; border: 1px solid #eee;'));
                    row.appendChild(makeTD(String(entry.parentCreator || ''), 'padding: 8px; color: #666; border: 1px solid #eee;'));

                    tbody.appendChild(row);
                }
                catch (e) {
                    this.log(`Skipping malformed PDF row: ${e}`);
                }
            }

            table.appendChild(tbody);
            container.appendChild(table);
            return pdfEntries.length;
        }
        catch (e) {
            let details = String(e?.message || e || 'Unknown error');
            this.log(`renderExportItems failed: ${details} :: ${e?.stack || ''}`);
            container.textContent = '';
            let warnWrap = (container.ownerDocument || window?.document).createElementNS('http://www.w3.org/1999/xhtml', 'div');
            warnWrap.setAttribute('style', 'padding: 20px; color: #6b4f00; background: #fff8e1; border: 1px solid #f3d27a; margin: 10px; border-radius: 4px;');

            let title = (container.ownerDocument || window?.document).createElementNS('http://www.w3.org/1999/xhtml', 'div');
            title.setAttribute('style', 'font-weight: 600; margin-bottom: 6px;');
            title.textContent = '无法完整读取当前视图条目，已自动降级为空列表。';

            let detail = (container.ownerDocument || window?.document).createElementNS('http://www.w3.org/1999/xhtml', 'div');
            detail.setAttribute('style', 'font-size: 12px; color: #7a5a00;');
            detail.textContent = details;

            warnWrap.appendChild(title);
            warnWrap.appendChild(detail);
            container.appendChild(warnWrap);
            return 0;
        }
    },

    async getCurrentVisiblePDFEntries(window) {
        let entryByAttachmentID = new Map();

        let addAttachment = async (attachment, parentItem = null) => {
            if (!attachment || typeof attachment.isAttachment !== 'function' || !attachment.isAttachment()) return;

            let attachmentID = attachment.id;
            if (attachmentID == null || entryByAttachmentID.has(attachmentID)) return;

            let filePath = '';
            let fileName = '';
            try {
                filePath = await this.getAttachmentFilePath(attachment);

                let parts = String(filePath).split(/[\\/]/);
                fileName = parts[parts.length - 1] || '';

                let attachmentFileName = String(attachment?.attachmentFilename || '');
                if (!fileName) {
                    fileName = attachmentFileName || String(attachment?.getField?.('title') || '');
                }

                // Detect PDF by metadata first, with runtime path/name as fallback.
                if (!this.isPDFAttachment(attachment, fileName || attachmentFileName, filePath)) {
                    return;
                }
            }
            catch (e) {
                this.log(`Attachment ${attachmentID}: file path resolution failed: ${e}`);
                let attachmentFileName = String(attachment?.attachmentFilename || '');
                fileName = attachmentFileName || String(attachment?.getField?.('title') || '');
                if (!this.isPDFAttachment(attachment, fileName, '')) {
                    return;
                }
            }

            if (!fileName) {
                fileName = 'Untitled.pdf';
                return;
            }

            let parentTitle = '';
            let parentCreator = '';
            if (parentItem && typeof parentItem.isRegularItem === 'function' && parentItem.isRegularItem()) {
                parentTitle = parentItem.getField('title') || '';
                parentCreator = parentItem.firstCreator || '';
            }

            entryByAttachmentID.set(attachmentID, {
                attachmentID,
                filePath,
                fileName,
                parentTitle,
                parentCreator,
            });
        };

        let regularItems = [];
        try {
            regularItems = this.getCurrentVisibleRegularItems(window);
        }
        catch (e) {
            this.log(`getCurrentVisibleRegularItems threw: ${e}`);
        }

        for (let item of regularItems) {
            if (!item || typeof item.getAttachments !== 'function') continue;

            let attIDs = [];
            try {
                attIDs = item.getAttachments() || [];
            }
            catch (e) {
                this.log(`getAttachments failed for item ${item.id}: ${e}`);
                continue;
            }

            for (let id of attIDs) {
                let att = this.safeGetItemByID(id, 'regularItem.getAttachments');
                await addAttachment(att, item);
            }
        }

        try {
            let selected = this.normalizeToArray(window?.ZoteroPane?.getSelectedItems?.());
            for (let item of selected) {
                if (!item) continue;

                if (typeof item.isAttachment === 'function' && item.isAttachment()) {
                    let parent = item.parentItemID ? this.safeGetItemByID(item.parentItemID, 'selectedAttachment.parentItemID') : null;
                    await addAttachment(item, parent);
                    continue;
                }

                if (typeof item.isRegularItem === 'function' && item.isRegularItem() && typeof item.getAttachments === 'function') {
                    for (let id of item.getAttachments() || []) {
                        let att = this.safeGetItemByID(id, 'selectedRegularItem.getAttachments');
                        await addAttachment(att, item);
                    }
                }
            }
        }
        catch (e) {
            this.log(`getSelectedItems fallback for PDF attachments failed: ${e}`);
        }

        try {
            let visibleAttachments = this.getCurrentVisibleAttachmentItems(window);
            for (let attachment of visibleAttachments) {
                if (!attachment) continue;
                let parent = attachment.parentItemID ? this.safeGetItemByID(attachment.parentItemID, 'visibleAttachment.parentItemID') : null;
                await addAttachment(attachment, parent);
            }
        }
        catch (e) {
            this.log(`getCurrentVisibleAttachmentItems failed: ${e}`);
        }

        let entries = Array.from(entryByAttachmentID.values());
        entries.sort((a, b) => {
            let ta = (a.parentTitle || '').toLowerCase();
            let tb = (b.parentTitle || '').toLowerCase();
            if (ta !== tb) return ta.localeCompare(tb);
            return (a.fileName || '').toLowerCase().localeCompare((b.fileName || '').toLowerCase());
        });
        return entries;
    },

    async getAttachmentFilePath(attachment) {
        if (!attachment) return '';

        try {
            if (typeof attachment.getFilePathAsync === 'function') {
                let path = await attachment.getFilePathAsync();
                if (path) return String(path);
            }
        }
        catch (e) {
            this.log(`getFilePathAsync failed for attachment ${attachment?.id}: ${e}`);
        }

        try {
            if (typeof attachment.getFilePath === 'function') {
                let path = attachment.getFilePath();
                if (path) return String(path);
            }
        }
        catch (e) {
            this.log(`getFilePath failed for attachment ${attachment?.id}: ${e}`);
        }

        try {
            if (typeof attachment.getFileAsync === 'function') {
                let file = await attachment.getFileAsync();
                let path = file?.path || file?.persistentDescriptor || '';
                if (path) return String(path);
            }
        }
        catch (e) {
            this.log(`getFileAsync failed for attachment ${attachment?.id}: ${e}`);
        }

        try {
            if (typeof attachment.getFile === 'function') {
                let file = attachment.getFile();
                let path = file?.path || file?.persistentDescriptor || '';
                if (path) return String(path);
            }
        }
        catch (e) {
            this.log(`getFile failed for attachment ${attachment?.id}: ${e}`);
        }

        return '';
    },

    async getAttachmentSourceFile(attachment, window) {
        if (!attachment) return null;

        try {
            if (typeof attachment.getFileAsync === 'function') {
                let file = await attachment.getFileAsync();
                let normalized = this.normalizeToNsIFile(file, window);
                if (normalized) return normalized;
            }
        }
        catch (e) {
            this.log(`getAttachmentSourceFile getFileAsync failed for ${attachment?.id}: ${e}`);
        }

        try {
            if (typeof attachment.getFile === 'function') {
                let file = attachment.getFile();
                let normalized = this.normalizeToNsIFile(file, window);
                if (normalized) return normalized;
            }
        }
        catch (e) {
            this.log(`getAttachmentSourceFile getFile failed for ${attachment?.id}: ${e}`);
        }

        let filePath = await this.getAttachmentFilePath(attachment);
        if (!filePath) return null;

        return this.nsIFileFromPathOrURI(filePath, window);
    },

    normalizeToNsIFile(fileLike, window) {
        if (!fileLike) return null;

        if (typeof fileLike.copyTo === 'function' && typeof fileLike.exists === 'function') {
            return fileLike;
        }

        let path = '';
        try {
            path = String(fileLike.path || fileLike.persistentDescriptor || '');
        }
        catch (e) {
            this.log(`normalizeToNsIFile path read failed: ${e}`);
            return null;
        }

        if (!path) return null;
        return this.nsIFileFromPathOrURI(path, window);
    },

    async copyFileToDirectoryUnique(srcFile, destDirFile, window) {
        if (!srcFile || !destDirFile) {
            throw new Error('copyFileToDirectoryUnique requires source and destination directory');
        }

        let srcLeafName = String(srcFile.leafName || 'attachment.pdf');
        let uniqueLeafName = this.getUniqueLeafName(destDirFile, srcLeafName);

        if (typeof srcFile.copyTo === 'function') {
            srcFile.copyTo(destDirFile, uniqueLeafName);
            return;
        }

        let srcPath = String(srcFile.path || '');
        if (!srcPath) {
            throw new Error('Source file path is unavailable');
        }

        let destFile = destDirFile.clone();
        destFile.append(uniqueLeafName);

        if (typeof IOUtils !== 'undefined' && typeof IOUtils.copy === 'function') {
            await IOUtils.copy(srcPath, destFile.path);
            return;
        }

        let componentsObj = this.getComponents(window);
        let Cc = componentsObj?.classes;
        let Ci = componentsObj?.interfaces;
        if (Cc && Ci && typeof Cc['@mozilla.org/network/file-input-stream;1'] !== 'undefined') {
            throw new Error('No supported file copy API available for this Zotero runtime');
        }

        throw new Error('No supported file copy API available');
    },

    getUniqueLeafName(destDirFile, preferredLeafName) {
        let baseName = String(preferredLeafName || 'attachment.pdf');
        let dotIndex = baseName.lastIndexOf('.');
        let stem = dotIndex > 0 ? baseName.slice(0, dotIndex) : baseName;
        let ext = dotIndex > 0 ? baseName.slice(dotIndex) : '';

        let candidate = baseName;
        let counter = 1;
        while (counter < 10000) {
            let probe = destDirFile.clone();
            probe.append(candidate);
            if (!probe.exists()) {
                return candidate;
            }
            candidate = `${stem} (${counter})${ext}`;
            counter++;
        }

        throw new Error('Unable to find a unique target filename');
    },

    nsIFileFromPathOrURI(filePath, window) {
        if (!filePath) return null;

        let componentsObj = this.getComponents(window);
        let Cc = componentsObj?.classes;
        let Ci = componentsObj?.interfaces;
        if (!Cc || !Ci) return null;

        if (/^file:\/\//i.test(filePath)) {
            try {
                if (typeof Services !== 'undefined' && Services?.io && typeof Services.io.newURI === 'function') {
                    let uri = Services.io.newURI(filePath);
                    if (uri && typeof uri.QueryInterface === 'function') {
                        let fileURL = uri.QueryInterface(Ci.nsIFileURL);
                        if (fileURL?.file) {
                            return fileURL.file;
                        }
                    }
                }
            }
            catch (e) {
                this.log(`file URI conversion failed for ${filePath}: ${e}`);
            }
        }

        try {
            let srcFile = Cc['@mozilla.org/file/local;1'].createInstance(Ci.nsIFile);
            srcFile.initWithPath(filePath);
            return srcFile;
        }
        catch (e) {
            this.log(`initWithPath failed for path=${filePath}: ${e}`);
            return null;
        }
    },

    isPDFAttachment(attachment, resolvedFileName = '', resolvedFilePath = '') {
        let ct = String(attachment?.attachmentContentType || '').toLowerCase();
        if (ct === 'application/pdf') return true;

        let fileName = String(attachment?.attachmentFilename || '').toLowerCase();
        if (fileName.endsWith('.pdf')) return true;

        let runtimeFileName = String(resolvedFileName || '').toLowerCase();
        if (runtimeFileName.endsWith('.pdf')) return true;

        let runtimeFilePath = String(resolvedFilePath || '').toLowerCase();
        if (runtimeFilePath.endsWith('.pdf')) return true;

        let title = String(attachment?.getField?.('title') || '').toLowerCase();
        return title.endsWith('.pdf');
    },

    safeGetItemByID(rawID, contextLabel = '') {
        let id = null;
        if (typeof rawID === 'number' && Number.isInteger(rawID)) {
            id = rawID;
        }
        else if (typeof rawID === 'string') {
            let parsed = Number.parseInt(rawID, 10);
            if (Number.isInteger(parsed)) {
                id = parsed;
            }
        }
        else if (rawID && typeof rawID === 'object' && Number.isInteger(rawID.id)) {
            id = rawID.id;
        }

        if (id == null) {
            if (rawID != null) {
                this.log(`Skipping invalid item id in ${contextLabel}: ${String(rawID)}`);
            }
            return null;
        }

        try {
            return Zotero.Items.get(id) || null;
        }
        catch (e) {
            this.log(`Zotero.Items.get(${id}) failed in ${contextLabel}: ${e}`);
            return null;
        }
    },

    normalizeToArray(candidate) {
        if (!candidate) return [];
        if (Array.isArray(candidate)) return candidate;

        try {
            if (typeof candidate[Symbol.iterator] === 'function') {
                return Array.from(candidate);
            }
        }
        catch (e) {
            this.log(`normalizeToArray iterator conversion failed: ${e}`);
        }

        let lengthValue = null;
        try {
            lengthValue = candidate.length;
        }
        catch (e) {
            this.log(`normalizeToArray length access failed: ${e}`);
            return [];
        }

        if (typeof lengthValue === 'number' && Number.isFinite(lengthValue) && lengthValue >= 0) {
            let out = [];
            let max = Math.min(Math.floor(lengthValue), 50000);
            for (let i = 0; i < max; i++) {
                try {
                    out.push(candidate[i]);
                }
                catch (e) {
                    this.log(`normalizeToArray index access failed at ${i}: ${e}`);
                    break;
                }
            }
            return out;
        }

        return [];
    },

    getCurrentVisibleRegularItems(window) {
        let pane = window?.ZoteroPane;
        if (!pane) return [];

        let items = [];
        let addRegularItems = (candidate) => {
            for (let item of this.normalizeToArray(candidate)) {
                if (item && typeof item.isRegularItem === 'function' && item.isRegularItem()) {
                    items.push(item);
                }
            }
        };

        let addItemsFromUnknownList = (candidate) => {
            candidate = this.normalizeToArray(candidate);
            if (!candidate.length) return;

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

                let item = this.safeGetItemByID(id, 'getCurrentVisibleRegularItems.addItemsFromUnknownList');
                if (item && typeof item.isRegularItem === 'function' && item.isRegularItem()) {
                    items.push(item);
                }
            }
        };

        try {
            let fromView = this.getCurrentVisibleItemsFromItemsView(window);
            for (let item of fromView) {
                if (item && typeof item.isRegularItem === 'function' && item.isRegularItem()) {
                    items.push(item);
                }
            }
        }
        catch (e) {
            this.log(`itemsView regular scan failed: ${e}`);
        }

        try {
            if (typeof pane.getSortedItems === 'function') {
                addItemsFromUnknownList(this.getPaneSortedItemsCompat(pane));
            }
        }
        catch (e) {
            this.log(`getSortedItems() failed: ${e}`);
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

    getCurrentVisibleAttachmentItems(window) {
        let pane = window?.ZoteroPane;
        if (!pane) return [];

        let attachments = [];
        let addAttachment = (item) => {
            if (item && typeof item.isAttachment === 'function' && item.isAttachment()) {
                attachments.push(item);
            }
        };

        let resolveItem = (raw) => {
            if (!raw) return null;
            if (typeof raw.isAttachment === 'function' || typeof raw.isRegularItem === 'function') {
                return raw;
            }

            let id = null;
            if (typeof raw === 'number' && Number.isInteger(raw)) {
                id = raw;
            }
            else if (typeof raw === 'string') {
                let parsed = Number.parseInt(raw, 10);
                if (Number.isInteger(parsed)) id = parsed;
            }
            else if (typeof raw === 'object' && Number.isInteger(raw.id)) {
                id = raw.id;
            }

            if (id == null) return null;
            return this.safeGetItemByID(id, 'getCurrentVisibleAttachmentItems.resolveItem');
        };

        let addFromUnknownList = (candidate) => {
            candidate = this.normalizeToArray(candidate);
            if (!candidate.length) return;
            for (let raw of candidate) {
                let item = resolveItem(raw);
                addAttachment(item);
            }
        };

        try {
            let fromView = this.getCurrentVisibleItemsFromItemsView(window);
            for (let item of fromView) {
                addAttachment(item);
            }
        }
        catch (e) {
            this.log(`itemsView direct attachment scan failed: ${e}`);
        }

        try {
            if (typeof pane.getSortedItems === 'function') {
                addFromUnknownList(this.getPaneSortedItemsCompat(pane));
            }
        }
        catch (e) {
            this.log(`getSortedItems() for attachments failed: ${e}`);
        }

        if (attachments.length === 0) {
            try {
                if (typeof pane.getSelectedItems === 'function') {
                    addFromUnknownList(pane.getSelectedItems() || []);
                }
            }
            catch (e) {
                this.log(`getSelectedItems() for attachments failed: ${e}`);
            }
        }

        if (attachments.length === 0) {
            try {
                let itemsView = pane.itemsView || pane._itemsView || pane.getItemsView?.();
                let rowCount = 0;
                if (typeof itemsView?.rowCount === 'number') {
                    rowCount = itemsView.rowCount;
                }
                else if (typeof itemsView?.getRowCount === 'function') {
                    rowCount = itemsView.getRowCount();
                }

                for (let i = 0; i < rowCount; i++) {
                    let row = null;
                    if (typeof itemsView.getRow === 'function') {
                        row = itemsView.getRow(i);
                    }
                    else if (typeof itemsView.getRowData === 'function') {
                        row = itemsView.getRowData(i);
                    }
                    else if (typeof itemsView.getItemAtIndex === 'function') {
                        row = itemsView.getItemAtIndex(i);
                    }

                    let candidate = row?.ref || row?.item || row;
                    let item = resolveItem(candidate);
                    addAttachment(item);
                }
            }
            catch (e) {
                this.log(`itemsView attachment scan failed: ${e}`);
            }
        }

        let dedup = new Map();
        for (let item of attachments) {
            if (item?.id != null) dedup.set(item.id, item);
        }
        return Array.from(dedup.values());
    },

    getCurrentVisibleItemsFromItemsView(window) {
        let pane = window?.ZoteroPane;
        if (!pane) return [];

        let itemsView = pane.itemsView || pane._itemsView || pane.getItemsView?.();
        if (!itemsView) return [];

        let out = [];
        let rowCount = 0;
        if (typeof itemsView.rowCount === 'number') {
            rowCount = itemsView.rowCount;
        }
        else if (typeof itemsView.getRowCount === 'function') {
            rowCount = itemsView.getRowCount();
        }

        for (let i = 0; i < rowCount; i++) {
            let row = null;
            try {
                if (typeof itemsView.getRow === 'function') {
                    row = itemsView.getRow(i);
                }
                else if (typeof itemsView.getRowData === 'function') {
                    row = itemsView.getRowData(i);
                }
                else if (typeof itemsView.getItemAtIndex === 'function') {
                    row = itemsView.getItemAtIndex(i);
                }
            }
            catch (e) {
                this.log(`itemsView row read failed at ${i}: ${e}`);
                continue;
            }

            let candidate = row?.ref || row?.item || row;
            let resolved = candidate;
            if (!resolved || (typeof resolved.isAttachment !== 'function' && typeof resolved.isRegularItem !== 'function')) {
                resolved = this.safeGetItemByID(candidate, 'getCurrentVisibleItemsFromItemsView.resolve');
            }

            if (resolved) out.push(resolved);
        }

        let dedup = new Map();
        for (let item of out) {
            if (item?.id != null) dedup.set(item.id, item);
        }
        return Array.from(dedup.values());
    },

    getPaneSortedItemsCompat(pane) {
        if (!pane || typeof pane.getSortedItems !== 'function') return [];

        let candidates = [
            () => pane.getSortedItems(),
            () => pane.getSortedItems(true),
            () => pane.getSortedItems(false),
            () => pane.getSortedItems({}),
        ];

        for (let getter of candidates) {
            try {
                let out = getter();
                let normalized = this.normalizeToArray(out);
                if (normalized.length) return normalized;
            }
            catch (e) {
                this.log(`getPaneSortedItemsCompat candidate failed: ${e}`);
            }
        }

        return [];
    },

    async pickFolderPath(window, defaultPath = '') {
        let componentsObj = this.getComponents(window);

        if (componentsObj?.classes && componentsObj?.interfaces) {
            let Ci = componentsObj.interfaces;
            let fp = componentsObj.classes['@mozilla.org/filepicker;1'].createInstance(Ci.nsIFilePicker);
            // Gecko APIs differ by version: some expect BrowsingContext, some expect Window.
            let pickerContext = window?.browsingContext || window;
            try {
                fp.init(pickerContext, 'Select Export Directory', Ci.nsIFilePicker.modeGetFolder);
            }
            catch (e) {
                this.log(`nsIFilePicker init with browsingContext/window failed, retry with window: ${e}`);
                fp.init(window, 'Select Export Directory', Ci.nsIFilePicker.modeGetFolder);
            }

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

            if (typeof fp.open === 'function') {
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

            if (typeof fp.show === 'function') {
                let rv = await fp.show();
                if (rv === Ci.nsIFilePicker.returnOK && fp.file) {
                    return fp.file.path;
                }
                return '';
            }

            this.log('nsIFilePicker has no open/show method in this context');
        }

        try {
            if (typeof ChromeUtils !== 'undefined' && typeof ChromeUtils.importESModule === 'function') {
                let { FilePicker } = ChromeUtils.importESModule('resource://gre/modules/FilePicker.sys.mjs');
                let picker = new FilePicker(window?.browsingContext || null);
                try {
                    picker.init(window?.browsingContext || window, 'Select Export Directory', picker.modeGetFolder);
                }
                catch (e) {
                    this.log(`FilePicker.sys.mjs init with browsingContext/window failed, retry with window: ${e}`);
                    picker.init(window, 'Select Export Directory', picker.modeGetFolder);
                }
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
