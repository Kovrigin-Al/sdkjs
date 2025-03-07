/*
 * (c) Copyright Ascensio System SIA 2010-2019
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(function(window, document){
    // Import
    var g_fontApplication = AscFonts.g_fontApplication;
    var CFontFileLoader = AscFonts.CFontFileLoader;
    var FONT_TYPE_EMBEDDED = AscFonts.FONT_TYPE_EMBEDDED;
    var CFontInfo = AscFonts.CFontInfo;
    var ImageLoadStatus = AscFonts.ImageLoadStatus;
    var CImage = AscFonts.CImage;

    function CGlobalFontLoader()
    {
        // сначала хотел писать "вытеснение" из этого мапа.
        // но тогда нужно хранить base64 строки. Это не круто. По памяти - даже
        // выигрыш будет. Не особо то шрифты жмутся lzw или deflate
        // поэтому лучше из памяти будем удалять base64 строки
        this.fonts_streams = [];

        // теперь вся информация о всех возможных шрифтах. Они во всех редакторах должны быть одни и те же
        this.fontFilesPath = "../../../../fonts/";
        this.fontFiles = AscFonts.g_font_files;
        this.fontInfos = AscFonts.g_font_infos;
        this.map_font_index = AscFonts.g_map_font_index;

        // теперь вся информация о всех встроенных шрифтах. Они должны удаляться при подгрузке нового файла
        this.embeddedFilesPath = "";
        this.embeddedFontFiles = [];
        this.embeddedFontInfos = [];

        // динамическая подгрузка шрифтов
        this.ThemeLoader = null;
        this.Api = null;
        this.fonts_loading = [];
        this.fonts_loading_after_style = [];
        this.bIsLoadDocumentFirst = false;
        this.currentInfoLoaded = null;

        this.loadFontCallBack     = null;
        this.loadFontCallBackArgs = null;

        this.IsLoadDocumentFonts2 = false;

        this.check_loaded_timer_id = -1;
        this.endLoadingCallback = null;

        this.perfStart = 0;

        this.put_Api = function(_api)
        {
            this.Api = _api;
        };
                
        this.LoadEmbeddedFonts = function(url, _fonts)
        {
            this.embeddedFilesPath = url;

            var _count = _fonts.length;

            if (0 == _count)
                return;

            this.embeddedFontInfos = new Array(_count);

            var map_files = {};
            for (var i = 0; i < _count; i++)
                map_files[_fonts[i].id] = _fonts[i].id;

            var index = 0;
            for (var i in map_files)
            {
                this.embeddedFontFiles[index] = new CFontFileLoader(map_files[i]);
                this.embeddedFontFiles[index].CanUseOriginalFormat = false;
                this.embeddedFontFiles[index].IsNeedAddJSToFontPath = false;
                map_files[i] = index++;
            }

            for (var i = 0; i < _count; i++)
            {
                var lStyle = 0;//_fonts[i].Style;
                if (0 == lStyle)
                    this.embeddedFontInfos[i] = new CFontInfo(_fonts[i].name, "", FONT_TYPE_EMBEDDED, map_files[_fonts[i].id], 0, -1, -1, -1, -1, -1, -1);
                else if (2 == lStyle)
                    this.embeddedFontInfos[i] = new CFontInfo(_fonts[i].name, "", FONT_TYPE_EMBEDDED, -1, -1, map_files[_fonts[i].id], _fonts[i].faceindex, -1, -1, -1, -1);
                else if (1 == lStyle)
                    this.embeddedFontInfos[i] = new CFontInfo(_fonts[i].name, "", FONT_TYPE_EMBEDDED, -1, -1, -1, -1, map_files[_fonts[i].id], _fonts[i].faceindex, -1, -1);
                else
                    this.embeddedFontInfos[i] = new CFontInfo(_fonts[i].name, "", FONT_TYPE_EMBEDDED, -1, -1, -1, -1, -1, -1, map_files[_fonts[i].id], _fonts[i].faceindex);
            }

            var _count_infos_ = this.fontInfos.length;
            for (var i = 0; i < _count; i++)
            {
                this.map_font_index[_fonts[i].name] = i + _count_infos_;
                this.fontInfos[i + _count_infos_] = this.embeddedFontInfos[i];
            }
        };

        this.SetStandartFonts = function()
        {
            //В стандартных шрифтах закоментированы те шрифты, которые были добавлены на docs.teamlab.com
            var standarts = window["standarts"];

            if (undefined == standarts)
            {
                standarts = [];
                for (var i = 0; i < this.fontInfos.length; i++)
                {
                    if (this.fontInfos[i].Name != "ASCW3")
                        standarts.push(this.fontInfos[i].Name);
                }
            }

            var _count = standarts.length;
            var _infos = this.fontInfos;
            var _map = this.map_font_index;
            for (var i = 0; i < _count; i++)
            {
                _infos[_map[standarts[i]]].Type = AscFonts.FONT_TYPE_STANDART;
            }
        };

        this.AddLoadFonts = function(name, need_styles)
        {
            var fontinfo = g_fontApplication.GetFontInfo(name);

            this.fonts_loading[this.fonts_loading.length] = fontinfo;
            this.fonts_loading[this.fonts_loading.length - 1].NeedStyles = (need_styles == undefined) ? 0x0F : need_styles;
			return fontinfo;
        };

        this.AddLoadFontsNotPick = function(info, need_styles)
        {
            this.fonts_loading[this.fonts_loading.length] = info;
            this.fonts_loading[this.fonts_loading.length - 1].NeedStyles = (need_styles == undefined) ? 0x0F : need_styles;
        };

        this.LoadDocumentFonts = function(_fonts, is_default)
        {
            if (this.IsLoadDocumentFonts2)
                return this.LoadDocumentFonts2(_fonts);

            // в конце метода нужно отдать список шрифтов
            var gui_fonts = [];
            var gui_count = 0;
            for (var i = 0; i < this.fontInfos.length; i++)
            {
                var info = this.fontInfos[i];
                if (AscFonts.FONT_TYPE_STANDART == info.Type)
                {
                    var __font = new AscFonts.CFont(info.Name, "", info.Type, info.Thumbnail);
                    gui_fonts[gui_count++] = __font;
                }
            }

            // сначала заполняем массив this.fonts_loading объекстами fontinfo
            for (var i in _fonts)
            {
                if (_fonts[i].type != FONT_TYPE_EMBEDDED)
                {
                    var info = this.AddLoadFonts(_fonts[i].name, _fonts[i].NeedStyles);

                    if (info.Type == AscFonts.FONT_TYPE_ADDITIONAL)
                    {
                        if (info.name != "ASCW3")
                        {
                            var __font = new AscFonts.CFont(info.Name, "", info.Type, info.Thumbnail);
                            gui_fonts[gui_count++] = __font;
                        }
                    }
                }
                else
                {
                    var ind = -1;
                    for (var j = 0; j < this.embeddedFontInfos.length; j++)
                    {
                        if (this.embeddedFontInfos[j].Name == _fonts[i].name)
                        {
                            this.AddLoadFontsNotPick(this.embeddedFontInfos[j], 0x0F);
                            break;
                        }
                    }
                }
            }

            this.Api.sync_InitEditorFonts(gui_fonts);

            // но только если редактор!!!
            if (this.Api.IsNeedDefaultFonts())
            {
                // теперь добавим шрифты, без которых редактор как без рук (спецсимволы + дефолтовые стили документа)
                this.AddLoadFonts("Arial", 0x0F);
                this.AddLoadFonts("Symbol", 0x0F);
                this.AddLoadFonts("Wingdings", 0x0F);
                this.AddLoadFonts("Courier New", 0x0F);
                this.AddLoadFonts("Times New Roman", 0x0F);
            }

            this.Api.asyncFontsDocumentStartLoaded();

            this.bIsLoadDocumentFirst = true;

            this.CheckFontsNeedLoadingLoad();
            this._LoadFonts();
        };

        this.CheckFontsNeedLoadingLoad = function()
        {
            var _fonts = this.fonts_loading;
            var _fonts_len = _fonts.length;

            var _need = false;
            for (var i = 0; i < _fonts_len; i++)
            {
                if (true == _fonts[i].CheckFontLoadStyles(this))
                    _need = true;
            }
            return _need;
        };

        this.CheckFontsNeedLoading = function(_fonts)
        {
            for (var i in _fonts)
            {
                var info = g_fontApplication.GetFontInfo(_fonts[i].name);
                var _isNeed = info.CheckFontLoadStylesNoLoad(this);
                if (_isNeed === true)
                    return true;
            }
            return false;
        };

        this.LoadDocumentFonts2 = function(_fonts, _blockType, _callback)
        {
            if (this.isWorking())
                return;

            this.endLoadingCallback = (undefined !== _callback) ? _callback : null;

            this.BlockOperationType = _blockType;
            // сначала заполняем массив this.fonts_loading объекстами fontinfo
            for (var i in _fonts)
            {
                this.AddLoadFonts(_fonts[i].name, 0x0F);
            }

            if (null == this.ThemeLoader)
                this.Api.asyncFontsDocumentStartLoaded(this.BlockOperationType);
            else
                this.ThemeLoader.asyncFontsStartLoaded();

            this.CheckFontsNeedLoadingLoad();
            this._LoadFonts();
        };

        var oThis = this;
        this._LoadFonts = function()
        {
            if (this.bIsLoadDocumentFirst === true && 0 === this.perfStart && this.fonts_loading.length > 0) {
                this.perfStart = performance.now();
            }
            if (0 == this.fonts_loading.length)
            {
                if (this.perfStart > 0) {
                    let perfEnd = performance.now();
                    AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onLoadFonts", perfEnd - this.perfStart), this.Api);
                    this.perfStart = 0;
                }
                if (null != this.endLoadingCallback)
                {
                    this.endLoadingCallback.call(this.Api);
                    this.endLoadingCallback = null;
                }
                else if (null == this.ThemeLoader)
                    this.Api.asyncFontsDocumentEndLoaded(this.BlockOperationType);
                else
                    this.ThemeLoader.asyncFontsEndLoaded();

                this.BlockOperationType = undefined;

                if (this.bIsLoadDocumentFirst === true)
                {
                    var _count = this.fonts_loading_after_style.length;
                    for (var i = 0; i < _count; i++)
                    {
                        var _info = this.fonts_loading_after_style[i];
                        _info.NeedStyles = 0x0F;
                        _info.CheckFontLoadStyles(this);
                    }
                    this.fonts_loading_after_style.splice(0, this.fonts_loading_after_style.length);

                    this.bIsLoadDocumentFirst = false;
                }
                return;
            }

            var fontinfo = this.fonts_loading[0];
            var IsNeed = fontinfo.CheckFontLoadStyles(this);

            if (IsNeed)
            {
                this.check_loaded_timer_id = setTimeout(oThis._check_loaded, 50);
            }
            else
            {
                if (this.bIsLoadDocumentFirst === true)
                {
                    this.Api.OpenDocumentProgress.CurrentFont++;
                    this.Api.SendOpenProgress();
                }

                this.fonts_loading_after_style[this.fonts_loading_after_style.length] = this.fonts_loading[0];
                this.fonts_loading.shift();
                this._LoadFonts();
            }
        };

        this.isWorking = function()
        {
            return (this.check_loaded_timer_id !== -1) ? true : false;
        };

        this._check_loaded = function()
        {
            oThis.check_loaded_timer_id = -1;
            if (0 == oThis.fonts_loading.length)
            {
                // значит асинхронно удалилось
                oThis._LoadFonts();
                return;
            }

            var current = oThis.fonts_loading[0];
            var IsNeed = current.CheckFontLoadStyles(oThis);
            if (true === IsNeed)
            {
                oThis.check_loaded_timer_id = setTimeout(oThis._check_loaded, 50);
            }
            else
            {
                if (oThis.bIsLoadDocumentFirst === true)
                {
                    oThis.Api.OpenDocumentProgress.CurrentFont++;
                    oThis.Api.SendOpenProgress();
                }
                
                oThis.fonts_loading_after_style[oThis.fonts_loading_after_style.length] = oThis.fonts_loading[0];
                oThis.fonts_loading.shift();
                oThis._LoadFonts();
            }
        };

        this.LoadFont = function(fontinfo, loadFontCallBack, loadFontCallBackArgs)
        {
            this.currentInfoLoaded = fontinfo;

            this.currentInfoLoaded = fontinfo;
            this.currentInfoLoaded.NeedStyles = 15; // все стили

            var IsNeed = this.currentInfoLoaded.CheckFontLoadStyles(this);

            if ( undefined === loadFontCallBack )
            {
                this.loadFontCallBack     = this.Api.asyncFontEndLoaded;
                this.loadFontCallBackArgs = this.currentInfoLoaded;
            }
            else
            {
                this.loadFontCallBack     = loadFontCallBack;
                this.loadFontCallBackArgs = loadFontCallBackArgs;
            }

            if (IsNeed)
            {
                this.Api.asyncFontStartLoaded();
                setTimeout(this.check_loaded, 20);
                return true;
            }
            else
            {
                this.currentInfoLoaded = null;
                return false;
            }
        };
        this.check_loaded = function()
        {
            var current = oThis.currentInfoLoaded;

            if (null == current)
                return;

            var IsNeed = current.CheckFontLoadStyles(oThis);
            if (IsNeed)
            {
                setTimeout(oThis.check_loaded, 50);
            }
            else
            {
                oThis.loadFontCallBack.call( oThis.Api, oThis.loadFontCallBackArgs );
                oThis.currentInfoLoaded = null;
            }
        };

        this.LoadFontsFromServer = function(_fonts)
        {
            var _count = _fonts.length;
            for (var i = 0; i < _count; i++)
            {
                var _info = g_fontApplication.GetFontInfo(_fonts[i]);
                if (undefined !== _info)
                {
                    _info.LoadFontsFromServer(this);
                }
            }
        }
    }
	CGlobalFontLoader.prototype.SetStreamIndexEmb = function(font_index, stream_index)
	{
		this.embeddedFontFiles[font_index].SetStreamIndex(stream_index);
	};
	
    function CGlobalImageLoader()
    {
        this.map_image_index = {};

        // loading
        this.Api = null;
        this.ThemeLoader = null;
        this.images_loading = null;

        this.bIsLoadDocumentFirst = false;

        this.bIsAsyncLoadDocumentImages = false;

        this.bIsLoadDocumentImagesNoByOrder = true;
        this.nNoByOrderCounter = 0;

        this.isBlockchainSupport = false;
        var oThis = this;

        if (window["AscDesktopEditor"] &&
            window["AscDesktopEditor"]["IsLocalFile"] &&
            window["AscDesktopEditor"]["isBlockchainSupport"])
        {
            this.isBlockchainSupport = (window["AscDesktopEditor"]["isBlockchainSupport"]() && !window["AscDesktopEditor"]["IsLocalFile"]());

            if (this.isBlockchainSupport)
            {
                Image.prototype.preload_crypto = function(_url)
                {
                    window["crypto_images_map"] = window["crypto_images_map"] || {};
                    if (!window["crypto_images_map"][_url])
                        window["crypto_images_map"][_url] = [];
                    window["crypto_images_map"][_url].push(this);

                    window["AscDesktopEditor"]["PreloadCryptoImage"](_url, AscCommon.g_oDocumentUrls.getLocal(_url));

                    oThis.Api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
                };

                Image.prototype["onload_crypto"] = function(_src, _crypto_data)
                {
                    if (_crypto_data && AscCommon.EncryptionWorker && AscCommon.EncryptionWorker.isCryptoImages())
                    {
                        AscCommon.EncryptionWorker.decryptImage(_src, this, _crypto_data);
                        return;
                    }
					this.crossOrigin = "";
                    this.src = _src;
                    oThis.Api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
                };
            }
        }

        this.put_Api = function(_api)
        {
            this.Api = _api;

            if (this.Api.IsAsyncOpenDocumentImages !== undefined)
            {
                this.bIsAsyncLoadDocumentImages = this.Api.IsAsyncOpenDocumentImages();
                if (this.bIsAsyncLoadDocumentImages)
                {
                    if (undefined === this.Api.asyncImageEndLoadedBackground)
                        this.bIsAsyncLoadDocumentImages = false;
                }
            }
        };
        
        this.LoadDocumentImages = function(_images, isCheckExists)
        {
            if (isCheckExists)
            {
                for (var i = _images.length - 1; i >= 0; i--)
                {
                    var _id = AscCommon.getFullImageSrc2(_images[i]);
                    if (this.map_image_index[_id] && (this.map_image_index[_id].Status === ImageLoadStatus.Complete))
                    {
                        _images.splice(i, 1);
                    }
                }

                if (0 === _images.length)
                    return;
            }

            // сначала заполним массив
            if (this.ThemeLoader == null)
                this.Api.asyncImagesDocumentStartLoaded();
            else
                this.ThemeLoader.asyncImagesStartLoaded();

            this.images_loading = [];

            for (var id in _images)
            {
                this.images_loading[this.images_loading.length] = AscCommon.getFullImageSrc2(_images[id]);
            }

            if (!this.bIsAsyncLoadDocumentImages)
            {
				this.nNoByOrderCounter = 0;
                this._LoadImages();
            }
            else
            {
                var _len = this.images_loading.length;
                for (var i = 0; i < _len; i++)
                {
                    this.LoadImageAsync(i);
                }

                this.images_loading.splice(0, _len);

                if (this.ThemeLoader == null)
                    this.Api.asyncImagesDocumentEndLoaded();
                else
                    this.ThemeLoader.asyncImagesEndLoaded();
            }
        };

        this.loadImageByUrl = function(_image, _url, isDisableCrypto)
        {
            if (this.isBlockchainSupport && (true !== isDisableCrypto))
                _image.preload_crypto(_url);
            else
                _image.src = _url;
        };

        this._LoadImages = function()
        {
			for (var i = 0; i < this.images_loading.length; i++)
            {
				var _id = this.images_loading[i];
				if (this.map_image_index[_id] && (this.map_image_index[_id].Status === ImageLoadStatus.Complete))
                {
                    this.images_loading.splice(i, 1);
                }
            }
			var _count_images = this.images_loading.length;

            if (0 == _count_images)
            {
				this.nNoByOrderCounter = 0;

                if (this.ThemeLoader == null)
                    this.Api.asyncImagesDocumentEndLoaded();
                else
                    this.ThemeLoader.asyncImagesEndLoaded();

                return;
            }

            for (var i = 0; i < _count_images; i++)
			{
				var _id = this.images_loading[i];
				var oImage = new CImage(_id);
				oImage.Status = ImageLoadStatus.Loading;
				oImage.Image = new Image();
				oThis.map_image_index[oImage.src] = oImage;
				oImage.Image.parentImage = oImage;
				oImage.Image.onload = function ()
				{
					this.parentImage.Status = ImageLoadStatus.Complete;
					oThis.nNoByOrderCounter++;

					if (oThis.bIsLoadDocumentFirst === true)
					{
						oThis.Api.OpenDocumentProgress.CurrentImage++;
						oThis.Api.SendOpenProgress();
					}

					if (!oThis.bIsLoadDocumentImagesNoByOrder)
					{
						oThis.images_loading.shift();
						oThis._LoadImages();
					}
					else if (oThis.nNoByOrderCounter == oThis.images_loading.length)
                    {
						oThis.images_loading = [];
						oThis._LoadImages();
                    }
				};
				oImage.Image.onerror = function ()
				{
					this.parentImage.Status = ImageLoadStatus.Complete;
					this.parentImage.Image = null;
					oThis.nNoByOrderCounter++;

					if (oThis.bIsLoadDocumentFirst === true)
					{
						oThis.Api.OpenDocumentProgress.CurrentImage++;
						oThis.Api.SendOpenProgress();
					}

					if (!oThis.bIsLoadDocumentImagesNoByOrder)
					{
						oThis.images_loading.shift();
						oThis._LoadImages();
					}
					else if (oThis.nNoByOrderCounter == oThis.images_loading.length)
					{
						oThis.images_loading = [];
						oThis._LoadImages();
					}
				};
				AscCommon.backoffOnErrorImg(oImage.Image, function(img) {
					oThis.loadImageByUrl(img, img.src);
				});
				//oImage.Image.crossOrigin = 'anonymous';
				oThis.loadImageByUrl(oImage.Image, oImage.src);

				if (!oThis.bIsLoadDocumentImagesNoByOrder)
                    return;
			}
        };

        this.LoadImage = function(src, Type)
        {
            var _image = this.map_image_index[src];
            if (undefined != _image)
                return _image;

            this.Api.asyncImageStartLoaded();

            var oImage = new CImage(src);
            oImage.Type = Type;
            oImage.Image = new Image();
            oImage.Status = ImageLoadStatus.Loading;
            oThis.map_image_index[oImage.src] = oImage;

            oImage.Image.onload = function(){
                oImage.Status = ImageLoadStatus.Complete;
                oThis.Api.asyncImageEndLoaded(oImage);
            };
            oImage.Image.onerror = function(){
                oImage.Image = null;
                oImage.Status = ImageLoadStatus.Complete;
                oThis.Api.asyncImageEndLoaded(oImage);
            };
            AscCommon.backoffOnErrorImg(oImage.Image, function(img) {
                oThis.loadImageByUrl(img, img.src);
            });
            //oImage.Image.crossOrigin = 'anonymous';
            this.loadImageByUrl(oImage.Image, oImage.src);
            return null;
        };

        this.LoadImageAsync = function(i)
        {
            var _id = oThis.images_loading[i];
            var oImage = new CImage(_id);
            oImage.Status = ImageLoadStatus.Loading;
            oImage.Image = new Image();
            oThis.map_image_index[oImage.src] = oImage;
            oImage.Image.onload = function(){
                oImage.Status = ImageLoadStatus.Complete;
                oThis.Api.asyncImageEndLoadedBackground(oImage);
            };
            oImage.Image.onerror = function(){
                oImage.Status = ImageLoadStatus.Complete;
                oImage.Image = null;
                oThis.Api.asyncImageEndLoadedBackground(oImage);
            };
            AscCommon.backoffOnErrorImg(oImage.Image, function(img) {
                oThis.loadImageByUrl(img, img.src);
            });
            //oImage.Image.crossOrigin = 'anonymous';
            oThis.loadImageByUrl(oImage.Image, oImage.src);
        };

        this.LoadImagesWithCallback = function(arr, loadImageCallBack, loadImageCallBackArgs, isDisableCrypto)
        {
            var arrAsync = [];
            var i = 0;
            for (i = 0; i < arr.length; i++)
            {
                if (this.map_image_index[arr[i]] === undefined)
                    arrAsync.push(arr[i]);
            }

            if (arrAsync.length == 0)
            {
				loadImageCallBack.call(this.Api, loadImageCallBackArgs);
                return;
            }

            let asyncImageCounter = arrAsync.length;
            const callback = loadImageCallBack.bind(this.Api, loadImageCallBackArgs);

			for (i = 0; i < arrAsync.length; i++)
			{
				var oImage = new CImage(arrAsync[i]);
				oImage.Image = new Image();
				oImage.Image.parentImage = oImage;
				oImage.Status = ImageLoadStatus.Loading;
				this.map_image_index[oImage.src] = oImage;

				oImage.Image.onload = function ()
				{
					this.parentImage.Status = ImageLoadStatus.Complete;
                    asyncImageCounter--;

					if (asyncImageCounter === 0)
					    callback();
				};
				oImage.Image.onerror = function ()
				{
					this.parentImage.Image = null;
					this.parentImage.Status = ImageLoadStatus.Complete;
                    asyncImageCounter--;

					if (asyncImageCounter === 0)
						callback();
				};
				AscCommon.backoffOnErrorImg(oImage.Image, function(img) {
					oThis.loadImageByUrl(img, img.src);
				});
				//oImage.Image.crossOrigin = 'anonymous';
                this.loadImageByUrl(oImage.Image, oImage.src, isDisableCrypto);
			}
        };
    }

    //---------------------------------------------------------export---------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].CGlobalFontLoader = CGlobalFontLoader;
    window['AscCommon'].g_font_loader = new CGlobalFontLoader();
    window['AscCommon'].g_image_loader = new CGlobalImageLoader();
})(window, window.document);
