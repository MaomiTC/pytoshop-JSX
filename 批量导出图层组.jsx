#target photoshop

// 导出图层组为JPG的脚本
(function() {
    // 确保有文档打开
    if (app.documents.length > 0) {
        try {
            var doc = app.activeDocument;
            
            // 检查文档是否已保存
            if (!doc.path) {
                alert("请先保存文档！");
                return;
            }
            
            var docPath = doc.path;
            
            // 创建导出文件夹
            var exportFolder = new Folder(docPath + "/exported_groups");
            if (!exportFolder.exists) {
                exportFolder.create();
            }
            
            // 检查是否有图层组
            if (doc.layerSets.length === 0) {
                alert("文档中没有图层组！");
                return;
            }
            
            var exportCount = 0;
            
            // 保存当前文档状态
            var docState = doc.activeHistoryState;
            
            // 获取所有顶层图层组
            var allLayerSets = [];
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].typename === "LayerSet") {
                    allLayerSets.push(doc.layers[i]);
                }
            }
            
            // 遍历所有图层组
            for (var i = 0; i < allLayerSets.length; i++) {
                var currentGroup = allLayerSets[i];
                
                try {
                    // 隐藏所有图层
                    for (var j = 0; j < doc.layers.length; j++) {
                        doc.layers[j].visible = false;
                    }
                    
                    // 只显示当前图层组
                    currentGroup.visible = true;
                    app.activeDocument = doc;
                    
                    // 导出为JPG
                    var fileName = currentGroup.name.replace(/[\/\\\:\*\?\"\<\>\|]/g, '_') + ".jpg";
                    var saveFile = new File(exportFolder + "/" + fileName);
                    
                    // 复制当前文档
                    var tempDoc = doc.duplicate();
                    
                    // JPG保存选项
                    var jpgSaveOptions = new JPEGSaveOptions();
                    jpgSaveOptions.formatOptions = FormatOptions.STANDARDBASELINE;
                    jpgSaveOptions.quality = 12; // 最高质量
                    jpgSaveOptions.embedColorProfile = true;
                    
                    // 保存文件
                    tempDoc.saveAs(saveFile, jpgSaveOptions, true, Extension.LOWERCASE);
                    
                    // 关闭临时文档
                    tempDoc.close(SaveOptions.DONOTSAVECHANGES);
                    
                    // 恢复到原始文档
                    app.activeDocument = doc;
                    
                    exportCount++;
                } catch(err) {
                    alert("导出图层组 '" + currentGroup.name + "' 时出错：" + err);
                }
            }
            
            // 恢复文档到原始状态
            doc.activeHistoryState = docState;
            
            if (exportCount > 0) {
                alert("导出完成！成功导出 " + exportCount + " 个图层组\n文件保存在: " + exportFolder);
            } else {
                alert("没有图层组可供导出！");
            }
            
        } catch(err) {
            alert("执行过程中出错：" + err);
        }
    } else {
        alert("请先打开一个文档！");
    }
})(); 