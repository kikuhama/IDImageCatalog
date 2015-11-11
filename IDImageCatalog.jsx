/*
  InDesignを使って，画像ファイルのカタログを作成する。
  Version 1.0.0
  ©Copyright Hamajima Shoten, Publishers & Kikuchi Ken 2015
*/

#target indesign

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

function dumpObj(obj) {
    $.writeln("----");
    $.writeln(obj.toString());
    for(var prop in obj) {
	try {
	    $.writeln("name: " + prop + "; value: " + obj[prop]);
	}
	catch(e) {
	    $.writeln("name: " + prop + "; cannot access this property.");
	}
    }
}

function max(a, b) {
    if(a > b) {
	return a;
    }
    else {
	return b;
    }
}

function min(a, b) {
    if(a < b) {
	return a;
    }
    else {
	return b;
    }
}

var OptionDialog = function() {
    var EDIT_HEIGHT = 20;
    var RADIO_HEIGHT = 20;
    var STATIC_OFFSET = 5;
    var STATIC_HEIGHT = 15;
    var BUTTON_HEIGHT = 20;
    var LINE_HEIGHT = 25;
    var y;
    var window = new Window("dialog", "読み込み設定");
    var option = {
	targetFolder: null,
	pageSize: "A4",
	pageMargin: 15
    };

    var checkCondition = function() {
	if(option.targetFolder && option.pageMargin > 0) {
	    window.footer.okButton.enabled = true;
	}
	else {
	    window.footer.okButton.enabled = false;
	}
    };

    var onSelectPaperSize = function() {
	if(this == window.paperGroup.a4Radio) {
	    option.pageSize = "A4";
	}
	else if(this == window.paperGroup.b5Radio) {
	    option.pageSize = "B5";
	}
    };
    
    var onSelectTarget = function() {
	if(this == window.targetGroup.activeDocRadio) {
	    // アクティブドキュメント
	    window.targetGroup.pathEdit.enabled = false;
	    window.targetGroup.selectFolderButton.enabled = false;
	    option.target = "activeDocument";
	}
	else if(this == window.targetGroup.filesInFolderRadio) {
	    // フォルダ内のファイル
	    window.targetGroup.pathEdit.enabled = true;
	    window.targetGroup.selectFolderButton.enabled = true;
	    option.target = "folder";
	}
	checkCondition();
    };

    var selectFolder = function() {
	var folder = Folder.selectDialog("処理対象フォルダの選択");
	if(folder) {
	    // フォルダが選択された
	    option.targetFolder = folder;
	    window.targetGroup.pathEdit.text = folder.fsName;
	}
	checkCondition();
    };

    var onPathChanging = function() {
	var path = this.text;
	var folder = new Folder(path);
	if(folder.exists) {
	    option.targetFolder = folder;
	}
	else {
	    option.targetFolder = null;
	}
	checkCondition();
    };

    var onMarginChanging = function() {
	var margin = Number(this.text);
	$.writeln(margin);
    };

    var _this = this;
    window.targetGroup = window.add("panel",
				    [10, 10, 400, 70],
				    "処理対象");
    y = 10;
    window.targetGroup.add("statictext",
			   [15, y + STATIC_OFFSET,
			    50, y + STATIC_OFFSET + STATIC_HEIGHT],
			   "フォルダ");
    window.targetGroup.pathEdit
	= window.targetGroup.add("edittext",
				 [55, y, 300, y + EDIT_HEIGHT]);
    window.targetGroup.pathEdit.onChanging = onPathChanging;
    window.targetGroup.selectFolderButton
	= window.targetGroup.add("button",
				 [305, y, 380, y + BUTTON_HEIGHT],
				 "参照");
    window.targetGroup.selectFolderButton.onClick = selectFolder;

    window.paperGroup = window.add("panel",
				    [10, 100, 400, 170],
				    "用紙");
    y = 10;
    window.paperGroup.a4Radio
	= window.paperGroup.add("radiobutton",
				 [15, y, 155, y + RADIO_HEIGHT],
				 "A4");
    window.paperGroup.a4Radio.value = true;
    window.paperGroup.a4Radio.onClick = onSelectPaperSize;
    window.paperGroup.b5Radio
	= window.paperGroup.add("radiobutton",
				 [160, y, 300, y + RADIO_HEIGHT],
				 "B5");
    window.paperGroup.b5Radio.onClick = onSelectPaperSize;
    y += RADIO_HEIGHT;
    window.paperGroup.add("statictext",
			  [15, y + STATIC_OFFSET,
			   70, y + STATIC_OFFSET + STATIC_HEIGHT],
			  "余白(mm)");
    window.paperGroup.marginEdit
	= window.paperGroup.add("edittext",
				[80, y, 130, y + EDIT_HEIGHT],
				option.pageMargin);
    window.paperGroup.marginEdit.onChanging = onMarginChanging;
    
    window.footer = window.add("panel", [10, 150, 400, 190]);
    y = 10;
    window.footer.cancelButton
	= window.footer.add("button",
			    [200, y, 280, y + BUTTON_HEIGHT],
			    "キャンセル");
    window.footer.cancelButton.onClick = function() {
	window.close();
    };
    window.footer.okButton
	= window.footer.add("button",
			    [290, y, 380, y + BUTTON_HEIGHT],
			    "OK");
    window.footer.okButton.onClick = function() {
	_this.ok = true;
	window.close();
    };
    checkCondition();
    
    this.window = window;
    this.option = option;
};

OptionDialog.prototype.show = function(handler) {
    this.ok = false;
    this.window.show();
    if(this.ok && handler) {
	handler();
    }
};

var CatalogCreator = function() {
    this.dX = 5;
    this.dY = 10;
    this.captionHeight = 4;
};

CatalogCreator.prototype.run = function(option) {
    this.doc = this.createDoc(option);
    this.currentPage = this.doc.pages[0];
    this.x0 = this.currentPage.bounds[1] + this.currentPage.marginPreferences.left;
    this.y0 = this.currentPage.bounds[0] + this.currentPage.marginPreferences.top;
    this.x1 = this.currentPage.bounds[3] - this.currentPage.marginPreferences.right;
    this.y1 = this.currentPage.bounds[2] - this.currentPage.marginPreferences.bottom;
    this.x = this.x0;
    this.y = this.y0;
    this.ymax = this.y;
    var imageFiles = this.getImageFiles(option.targetFolder);
    var i;
    var imageRect;
    for(i=0; i<imageFiles.length; ++i) {
	imageRect = this.placeImage(imageFiles[i]);
	if(imageRect) {
	    this.createCaption(imageFiles[i], imageRect);
	}
    }
};

CatalogCreator.prototype.createCaption = function(imageFile, imageRect) {
    var textFrame = this.currentPage.textFrames.add();
    var bounds = textFrame.geometricBounds;
    var imageBounds = imageRect.geometricBounds;
    bounds[0] = imageBounds[2];
    bounds[1] = imageBounds[1];
    bounds[2] = bounds[0] + this.captionHeight;
    bounds[3] = imageBounds[3];
    textFrame.geometricBounds = bounds;
    textFrame.contents = imageFile.displayName;
};

CatalogCreator.prototype.placeImage = function(imageFile) {
    this.currentPage.place(imageFile, [this.x, this.y]);
    var rect = this.findImageRect(imageFile);
    if(rect) {
	var bounds = rect.geometricBounds;
	if(this.x != this.x0 && bounds[3] > this.x1) { // 横方向あふれ
	    this.x = this.x0;
	    this.y = this.ymax + this.dY;
	    rect.move([this.x, this.y]);
	    bounds = rect.geometricBounds;
	}
	this.x = bounds[3] + this.dX;
	this.ymax = max(this.ymax, bounds[2]);
	
	if(this.y != this.y0 && bounds[2] > this.y1) { // 縦方向あふれ
	    this.newPage();
	    rect.remove();
	    rect = this.placeImage(imageFile);
	}
    }

    return rect;
};

CatalogCreator.prototype.newPage = function() {
    this.currentPage = this.doc.pages.add(LocationOptions.atEnd);
    this.x0 = this.currentPage.bounds[1] + this.currentPage.marginPreferences.left;
    this.y0 = this.currentPage.bounds[0] + this.currentPage.marginPreferences.top;
    this.x1 = this.currentPage.bounds[3] - this.currentPage.marginPreferences.right;
    this.y1 = this.currentPage.bounds[2] - this.currentPage.marginPreferences.bottom;
    this.x = this.x0;
    this.y = this.y0;
    this.ymax = this.y;
    return this.currentPage;
};

CatalogCreator.prototype.findImageRect = function(imageFile) {
    var i;
    var rect;
    var r;
    var g;
    for(i=0; i<this.currentPage.rectangles.length; ++i) {
	r = this.currentPage.rectangles[i];
	if(r.graphics.length > 0) {
	    g = r.graphics[0];
	    if(g && g.itemLink && g.itemLink.name == imageFile.displayName) {
		rect = r;
	    }
	}
    }
    return rect;
};

CatalogCreator.prototype.createDoc = function(option) {
    var docPreset = app.documentPresets[0];
    docPreset.pageSize = option.pageSize;
    docPreset.facingPages = false;
    docPreset.top = docPreset.left = docPreset.bottom = docPreset.right = option.pageMargin;
    var doc = app.documents.add(true, docPreset);
    return doc;
};

CatalogCreator.prototype.getImageFiles = function(folder) {
    var imageFiles = [];
    var files = folder.getFiles("*");
    var i;
    for(i=0; i<files.length; ++i) {
	var f = files[i];
	if(f instanceof Folder) {
	    imageFiles.concat(this.getImageFiles(f));
	}
	else if(this.isImageFile(f)) {
	    imageFiles.push(f);
	}
    }
    return imageFiles;
};

CatalogCreator.prototype.isImageFile = function(file) {
    return(file.name.match(/\.(eps|ai|pdf|psd|jpg|tif|png)$/i)
	   && file.name.charAt(0) != ".");
};

var optDialog = new OptionDialog();
optDialog.show(function() {
    var creator = new CatalogCreator();
    creator.run(optDialog.option);
});

