/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
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

(function (window)
{

	function CCellFrameManager(oApi)
	{
		this.api = oApi;
		this.generalDocumentUrls = {};
	}

	CCellFrameManager.prototype.getWorkbookBinary = function ()
	{
		const oBinaryFileWriter = new AscCommonExcel.BinaryFileWriter(this.api.wbModel);
		const arrBinaryData = oBinaryFileWriter.Write().split(';');
		return arrBinaryData[arrBinaryData.length - 1];
	}
	CCellFrameManager.prototype.getAllImageIds = function ()
	{
		const arrRasterImageIds = [];
		const arrWorksheetLength = this.api.wbModel.aWorksheets.length;
		for (let i = 0; i < arrWorksheetLength; i += 1)
		{
			const oWorksheet = this.api.wbModel.aWorksheets[i];
			const arrDrawings = oWorksheet.Drawings;
			if (arrDrawings)
			{
				for (let j = 0; j < arrDrawings.length; j += 1)
				{
					const oDrawing = arrDrawings[j];
					oDrawing.graphicObject.getAllRasterImages(arrRasterImageIds);
				}
			}
		}
		return arrRasterImageIds;
	}

	CCellFrameManager.prototype.getChartObject = function ()
	{

	};
	CCellFrameManager.prototype.getImagesForHistory = function ()
	{
		const arrRasterImageIds = this.getAllImageIds();
		const urlsForAddToHistory = [];
		for (let i = 0; i < arrRasterImageIds.length; i += 1)
		{
			const url = AscCommon.g_oDocumentUrls.mediaPrefix + arrRasterImageIds[i];
			if (!(this.generalDocumentUrls[url] && this.generalDocumentUrls[url] === AscCommon.g_oDocumentUrls.getUrls()[url]))
			{
				urlsForAddToHistory.push(arrRasterImageIds[i]);
			}
		}

		return urlsForAddToHistory;
	}

	CCellFrameManager.prototype.obtain = function (oInfo)
	{
		const sStream = oInfo["binary"];
		this.setGeneralDocumentUrls(oInfo["documentImageUrls"]);
		this.openWorkbookData(sStream);
	};
	CCellFrameManager.prototype.setGeneralDocumentUrls = function (oPr)
	{
		this.generalDocumentUrls = oPr;
	};
	CCellFrameManager.prototype.getGeneralImageUrl = function (sImageId)
	{
		return this.generalDocumentUrls[sImageId];
	};
	CCellFrameManager.prototype.openWorkbookData = function (sStream)
	{
		const oFile = new AscCommon.OpenFileResult();
		oFile.bSerFormat = AscCommon.checkStreamSignature(sStream, AscCommon.c_oSerFormat.Signature);
		oFile.data = sStream;
		this.api.asc_CloseFile();

		this.api.imagesFromGeneralEditor = this.generalDocumentUrls;
		this.api.openDocument(oFile);
	}
	CCellFrameManager.prototype.isDiagramEditor = function ()
	{
		return false;
	};
	CCellFrameManager.prototype.isOleEditor = function ()
	{
		return false;
	};
	CCellFrameManager.prototype.updateGeneralDiagramCache = function (aRanges)
	{

	}

	CCellFrameManager.prototype.sendLoadImages = function (arrImages, token, bNotShowError)
	{
		this.sendFromFrameToGeneralEditor(new CFrameImageData(arrImages, token, bNotShowError));
	}

	CCellFrameManager.prototype.sendFromFrameToGeneralEditor = function (oSendObject)
	{
		this.api.sendFromFrameToGeneralEditor(oSendObject);
	};

	CCellFrameManager.prototype.sendFromGeneralToFrameEditor = function (oSendObject)
	{
		this.api.sendFromGeneralToFrameEditor(oSendObject);
	};

	function COleCellFrameManager(oApi)
	{
		CCellFrameManager.call(this, oApi);
		this.imageWidthCoefficient = null;
		this.imageHeightCoefficient = null;
		this.isFromSheetEditor = false;
	}

	AscFormat.InitClassWithoutType(COleCellFrameManager, CCellFrameManager);

	COleCellFrameManager.prototype.getBase64Image = function ()
	{
		return this.api.wb.getImageFromTableOleObject();
	};
	COleCellFrameManager.prototype.getImageWidthCoefficient = function ()
	{
		return this.imageWidthCoefficient || undefined;
	}
	COleCellFrameManager.prototype.getImageHeightCoefficient = function ()
	{
		return this.imageHeightCoefficient || undefined;
	}
	COleCellFrameManager.prototype.getBinary = function ()
	{
		return {
			'binary'               : this.getWorkbookBinary(),
			'base64Image'          : this.getBase64Image(),
			'imagesForAddToHistory': this.getImagesForHistory(),
			'widthCoefficient'     : this.getImageWidthCoefficient(),
			'heightCoefficient'    : this.getImageHeightCoefficient()
		};
	};
	COleCellFrameManager.prototype.calculateImageSaveCoefficients = function (nImageHeight, nImageWidth)
	{
		if (!nImageHeight || !nImageWidth)
			return;

		const saveImageCoefficients = this.api.getScaleCoefficientsForOleTableImage(nImageWidth, nImageHeight)
		this.imageWidthCoefficient = saveImageCoefficients.widthCoefficient;
		this.imageHeightCoefficient = saveImageCoefficients.heightCoefficient;
	}
	COleCellFrameManager.prototype.obtain = function (oInfo)
	{
		this.isFromSheetEditor = oInfo["isFromSheetEditor"];
		this.calculateImageSaveCoefficients(oInfo["imageHeight"], oInfo["imageWidth"]);
		CCellFrameManager.prototype.obtain.call(this, oInfo);
	}
	COleCellFrameManager.prototype.isFromSheetEditor = function ()
	{
		return this.isFromSheetEditor;
	}

	COleCellFrameManager.prototype.setAfterLoadCallback = function ()
	{

		const oApi = this.api;
		const oThis = this;
		oApi.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);
		// на случай, если изображение поставили на загрузку, закрыли редактор, и потом опять открыли
		oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);

		this.sendFromFrameToGeneralEditor({
			"type": AscCommon.c_oAscFrameDataType.OpenFrame
		});

		oApi.fAfterLoad = function ()
		{
			oThis.api.wb.scrollToOleSize();
			// добавляем первый поинт после загрузки, чтобы в локальную историю добавился либо стандартный oleSize, либо заданный пользователем
			const oleSize = oThis.api.wb.getOleSize();
			oleSize.addPointToLocalHistory();

			oThis.api.wb.onFrameEditorReady();
			oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);
			delete oApi.fAfterLoad;
		}
	};
	COleCellFrameManager.prototype.openWorkbookData = function (sStream)
	{
		this.setAfterLoadCallback();
		CCellFrameManager.prototype.openWorkbookData.call(this, sStream);
	}

	COleCellFrameManager.prototype.isOleEditor = function ()
	{
		return true;
	};


	function CDiagramCellFrameManager(oApi)
	{
		CCellFrameManager.call(this, oApi);
		this.arrAfterLoadCallbacks = [];
		this.mainDiagram = null;
	}
	AscFormat.InitClassWithoutType(CDiagramCellFrameManager, CCellFrameManager);

	CDiagramCellFrameManager.prototype.isDiagramEditor = function ()
	{
		return true;
	};
	CDiagramCellFrameManager.prototype.repairDiagramXLSX = function ()
	{
		const oThis = this;
		AscFormat.ExecuteNoHistory(function ()
		{
			AscCommonExcel.executeInR1C1Mode(false,
				function ()
				{
					oThis.fillWorkbookFromDiagramCache();
				});
		}, this, []);
	}
	CDiagramCellFrameManager.prototype.fillWorkbookFromDiagramCache = function ()
	{
		let max_r = 0, max_c = 0;
		const parserHelp = AscCommon.parserHelp;
		const series = this.mainDiagram.getAllSeries();
		let ser;
		const worksheet = this.api.wb.getWorksheet();
		const model = worksheet.model;
		const oThis = this;
		function fFillCell(oCell, sNumFormat, value)
		{
			var oCellValue = new AscCommonExcel.CCellValue();
			if (AscFormat.isRealNumber(value))
			{
				oCellValue.number = value;
				oCellValue.type = AscCommon.CellValueType.Number;
			}
			else
			{
				oCellValue.text = value;
				oCellValue.type = AscCommon.CellValueType.String;
			}
			oCell.setNumFormat(sNumFormat);
			oCell.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, oCellValue));
		}

		function fillTableFromRef(ref)
		{
			var cache = ref.numCache ? ref.numCache : (ref.strCache ? ref.strCache : null);
			var lit_format_code;
			if (cache)
			{
				lit_format_code = (typeof cache.formatCode === "string" && cache.formatCode.length > 0) ? cache.formatCode : "General";

				var sFormula = ref.f + "";
				if (sFormula[0] === '(')
					sFormula = sFormula.slice(1);
				if (sFormula[sFormula.length - 1] === ')')
					sFormula = sFormula.slice(0, -1);
				var f1 = sFormula;

				var arr_f = f1.split(",");
				var pt_index = 0, i, j, pt, nPtCount, k;
				for (i = 0; i < arr_f.length; ++i)
				{
					var parsed_ref = parserHelp.parse3DRef(arr_f[i]);
					if (parsed_ref)
					{
						var source_worksheet = oThis.api.wbModel.getWorksheetByName(parsed_ref.sheet);
						if (source_worksheet === model)
						{
							var range = source_worksheet.getRange2(parsed_ref.range);
							if (range)
							{
								range = range.bbox;

								if (range.r1 > max_r)
									max_r = range.r1;
								if (range.r2 > max_r)
									max_r = range.r2;

								if (range.c1 > max_c)
									max_c = range.c1;
								if (range.c2 > max_c)
									max_c = range.c2;

								if (i === arr_f.length - 1)
								{
									nPtCount = cache.getPtCount();
									if ((nPtCount - pt_index) <= (range.r2 - range.r1 + 1))
									{
										for (k = range.c1; k <= range.c2; ++k)
										{
											for (j = range.r1; j <= range.r2; ++j)
											{
												source_worksheet._getCell(j, k, function (cell)
												{
													pt = cache.getPtByIndex(pt_index + j - range.r1);
													if (pt)
													{
														fFillCell(cell, typeof pt.formatCode === "string" && pt.formatCode.length > 0 ? pt.formatCode : lit_format_code, pt.val);
													}
												});
											}
										}
										pt_index += (range.r2 - range.r1 + 1);
									}
									else if ((nPtCount - pt_index) <= (range.c2 - range.c1 + 1))
									{
										for (k = range.r1; k <= range.r2; ++k)
										{
											for (j = range.c1; j <= range.c2; ++j)
											{
												source_worksheet._getCell(k, j, function (cell)
												{
													pt = cache.getPtByIndex(pt_index + j - range.c1);
													if (pt)
													{
														fFillCell(cell, typeof pt.formatCode === "string" && pt.formatCode.length > 0 ? pt.formatCode : lit_format_code, pt.val);
													}
												});
											}
										}
										pt_index += (range.c2 - range.c1 + 1);
									}
								}
								else
								{
									if (range.r1 === range.r2)
									{
										for (j = range.c1; j <= range.c2; ++j)
										{
											source_worksheet._getCell(range.r1, j, function (cell)
											{
												pt = cache.getPtByIndex(pt_index);
												if (pt)
												{
													fFillCell(cell, typeof pt.formatCode === "string" && pt.formatCode.length > 0 ? pt.formatCode : lit_format_code, pt.val);
												}
												++pt_index;
											});
										}
									}
									else
									{
										for (j = range.r1; j <= range.r2; ++j)
										{
											source_worksheet._getCell(j, range.c1, function (cell)
											{
												pt = cache.getPtByIndex(pt_index);
												if (pt)
												{
													fFillCell(cell, typeof pt.formatCode === "string" && pt.formatCode.length > 0 ? pt.formatCode : lit_format_code, pt.val);
												}
												++pt_index;
											});
										}
									}
								}
							}
						}
					}
				}
			}
		}

		var first_num_ref;
		if (series[0])
		{
			if (series[0].val)
				first_num_ref = series[0].val.numRef;
			else if (series[0].yVal)
				first_num_ref = series[0].yVal.numRef;
		}
		if (first_num_ref)
		{
			var resultRef = parserHelp.parse3DRef(first_num_ref.f);
			if (resultRef)
			{
				model.workbook.aWorksheets[0].sName = resultRef.sheet;
				var oCat, oVal;
				for (var i = 0; i < series.length; ++i)
				{
					ser = series[i];
					oVal = ser.val || ser.yVal;
					if (oVal && oVal.numRef)
					{
						fillTableFromRef(oVal.numRef);
					}
					oCat = ser.cat || ser.xVal;
					if (oCat)
					{
						if (oCat.numRef)
						{
							fillTableFromRef(oCat.numRef);
						}
						if (oCat.strRef)
						{
							fillTableFromRef(oCat.strRef);
						}
					}
					if (ser.tx && ser.tx.strRef)
					{
						fillTableFromRef(ser.tx.strRef);
					}
				}
			}
		}

		worksheet._updateRange(new Asc.Range(0, 0, max_c, max_r));
		worksheet.draw();
	}
	CDiagramCellFrameManager.prototype.setMainDiagram = function (oInfo)
	{
		AscFormat.ExecuteNoHistory(function ()
		{
			const asc_chart_binary = new Asc.asc_CChartBinary();
			asc_chart_binary.asc_setBinary(oInfo["binary"]);
			const oModel = this.api.wb.getWorksheet().model;
			const oNewChartSpace = asc_chart_binary.getChartSpace(oModel);
			oNewChartSpace.setWorksheet(oModel);

			oNewChartSpace.setFrameChart(true);
			this.mainDiagram = oNewChartSpace;
		}, this, []);
	}
	CDiagramCellFrameManager.prototype.preObtain = function (oInfo)
	{
		this.setAfterLoadCallback(oInfo);
		if (oInfo["workbookBinary"])
		{
			this.obtain({"binary": oInfo["workbookBinary"], "documentImageUrls": oInfo["documentImageUrls"]});
		}
		else
		{
			this.obtainWithRepair({"binary": AscCommon.getEmpty(), "documentImageUrls": oInfo["documentImageUrls"]});
		}
	}
	CDiagramCellFrameManager.prototype.obtainWithRepair = function (oInfo)
	{
		this.setRepairAfterLoadCallback();
		this.obtain(oInfo);
	}
	CDiagramCellFrameManager.prototype.setRepairAfterLoadCallback = function ()
	{
		this.arrAfterLoadCallbacks.push(this.repairDiagramXLSX.bind(this));
	}
	CDiagramCellFrameManager.prototype.setAfterLoadCallback = function (oInfo)
	{
		const oApi = this.api;
		const oThis = this;
		oApi.fAfterLoad = function ()
		{
			oThis.setMainDiagram(oInfo);
			for (let i = 0; i < oThis.arrAfterLoadCallbacks.length; i += 1)
			{
				oThis.arrAfterLoadCallbacks[i]();
			}
			oThis.api.wb.onFrameEditorReady();
			oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);
			delete oApi.fAfterLoad;
		}
	}

	CDiagramCellFrameManager.prototype.getBinary = function ()
	{
		const oDiagramBinary = new Asc.asc_CChartBinary(this.mainDiagram);
		oDiagramBinary["workbookBinary"] = this.getWorkbookBinary();
		oDiagramBinary["imagesForAddToHistory"] = this.getImagesForHistory();
		return oDiagramBinary;
	}

	CDiagramCellFrameManager.prototype.updateGeneralDiagramCache = function (aRanges)
	{
		const aRefsToChange = [];
		this.mainDiagram.collectIntersectionRefs(aRanges, aRefsToChange);
		for (let i = 0; i < aRefsToChange.length; i += 1)
		{
			aRefsToChange[i].updateCacheAndCat();
		}
		this.sendUpdateDiagram();
	};
	CDiagramCellFrameManager.prototype.sendUpdateDiagram = function ()
	{
		this.sendFromFrameToGeneralEditor(new CFrameUpdateDiagramData(this.mainDiagram, true));
	};

	CDiagramCellFrameManager.prototype.getChartObject = function ()
	{
		const oGraphicController = this.api.getGraphicController();
		if (oGraphicController)
		{
			const oProps = oGraphicController.getPropsFromChart(this.mainDiagram);
			oProps.setFUpdateGeneralChart(this.sendUpdateDiagram.bind(this));
			return oProps;
		}
	};

	function CFrameUpdateDiagramData(oDiagram, bNoHistory)
	{
		const oBinary = new Asc.asc_CChartBinary(oDiagram);
		const oData = {"binary": oBinary, "noHistory": !!bNoHistory};
		CFrameData.call(this, AscCommon.c_oAscFrameDataType.UpdateDiagramInGeneral, oData);
	}

	function CGeneralUpdateDiagramData(oData)
	{
		CFrameData.call(this, AscCommon.c_oAscFrameDataType.UpdateDiagramInFrame, oData);
	}


	function CFrameData(type, information)
	{
		this["information"] = information;
		this["type"] = type;
	}

	function CFrameImageData(arrImages, token, bNotShowError)
	{
		const oData = {
			"images"       : arrImages,
			"token"        : token,
			"bNotShowError": bNotShowError
		}
		CFrameData.call(this, AscCommon.c_oAscFrameDataType.SendImageUrls, oData);
	}


	function CFrameBinaryLoader(oApi)
	{
		this.api = oApi;
		this.binary = null;
	}
	CFrameBinaryLoader.prototype.loadFrame = function ()
	{

	}
	CFrameBinaryLoader.prototype.isOpenedFrame = function ()
	{
		return true;
	}
	CFrameBinaryLoader.prototype.isOpenedChartFrame = function ()
	{
		return false;
	};
	CFrameBinaryLoader.prototype.isOpenedOleFrame = function ()
	{
		return false;
	};
	CFrameBinaryLoader.prototype.destroy = function ()
	{
		this.api.frameLoader = null;
	}


	function CFrameDiagramBinaryLoader(oApi, oChart)
	{
		CFrameBinaryLoader.call(this, oApi);
		this.chart = oChart;
		this.canLoad = true;
		this.XLSXBase64 = null;
	}
	AscFormat.InitClassWithoutType(CFrameDiagramBinaryLoader, CFrameBinaryLoader);

	CFrameDiagramBinaryLoader.getBase64 = function (arrStream)
	{
		const nDataSize = arrStream.length;
		const sData = AscCommon.Base64.encode(arrStream);
		return "XLSY;v2;" + nDataSize + ";" + sData;
	};

	CFrameDiagramBinaryLoader.prototype.createChartSpace = function (nType, oPlaceholder)
	{

	};
	CFrameDiagramBinaryLoader.prototype.isOpenedChartFrame = function ()
	{
		return true;
	};

	CFrameDiagramBinaryLoader.prototype.canLoadFrame = function ()
	{
		return this.canLoad;
	};
	CFrameDiagramBinaryLoader.prototype.getBinaryChart = function ()
	{
		const oBinaryChart = new Asc.asc_CChartBinary(this.chart);
		oBinaryChart.setWorkbookBinary(this.XLSXBase64);
		return oBinaryChart;

	};
	CFrameDiagramBinaryLoader.prototype.setXLSX = function (sStream)
	{
		if (sStream && sStream.length)
		{
			if (typeof sStream === 'string')
			{
				this.XLSXBase64 = sStream;
			}
			else
			{
				this.XLSXBase64 = Array.from(sStream);
			}
		}
		else
		{
			this.XLSXBase64 = null;
		}
	};
	CFrameDiagramBinaryLoader.prototype.resolvePromise = function (sStream)
	{
		if (this.isTruthStream(sStream))
		{
			this.setXLSX(sStream);
			this.loadFrame();
		}
		else
		{
			this.rejectPromise();
		}
	};
	CFrameDiagramBinaryLoader.prototype.rejectPromise = function ()
	{

	}
	CFrameDiagramBinaryLoader.prototype.setCanLoad = function (bPr)
	{
		this.canLoad = bPr;
	}
	CFrameDiagramBinaryLoader.prototype.isExternal = function ()
	{
		return this.chart.isExternal()
	}
	CFrameDiagramBinaryLoader.prototype.isTruthStream = function (arrStream)
	{
		return (arrStream && (arrStream.length !== 0)) || !this.isExternal();
	}
	CFrameDiagramBinaryLoader.prototype.startLoadWorksheet = function ()
	{
		this.api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
	};
	CFrameDiagramBinaryLoader.prototype.endLoadWorksheet = function ()
	{
		this.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
	};
	CFrameDiagramBinaryLoader.prototype.tryOpen = function ()
	{
		this.startLoadWorksheet();
		const oPromise = this.getPromise();
		oPromise.then(this.resolvePromise.bind(this));
	};
	CFrameDiagramBinaryLoader.prototype.loadFrame = function ()
	{
		this.api.asc_onOpenChartFrame();
		this.endLoadWorksheet();
		this.api.sendEvent('asc_doubleClickOnChart', this.getBinaryChart());
	};

	CFrameDiagramBinaryLoader.prototype.getPromise = function ()
	{
		if (this.isExternal())
		{
			return this.getExternalPromise();
		}
		return this.getNestedPromise();
	};
	CFrameDiagramBinaryLoader.prototype.getExternalPromise = function ()
	{
		const oExternalDataChartManager = new CFrameDiagramExternalDataManager(this.chart, this.api);
		return oExternalDataChartManager.getPromise();
	}
	CFrameDiagramBinaryLoader.prototype.getNestedPromise = function ()
	{
		const oThis = this;
		return new Promise(function (resolve)
		{
			resolve(oThis.chart.XLSX.length ? CFrameDiagramBinaryLoader.getBase64(oThis.chart.XLSX) : null);
		});
	}

	function CExternalDataLoader(arrExternalReference, oApi, fCallback)
	{
		this.externalReferences = arrExternalReference || [];
		this.api = oApi;
		this.isLocalDesktop = window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]();
		this.fCallback = fCallback;
	}

	CExternalDataLoader.prototype.updateExternalData = function ()
	{
		if (this.externalReferences)
		{
			if (this.isLocalDesktop)
			{
				this.resolveUpdateData();
			}
			else
			{
				this.api.sendEvent("asc_onUpdateExternalReference", this.externalReferences, this.resolveUpdateData.bind(this));
			}
		}
	};

	CExternalDataLoader.prototype.resolveUpdateData = function (arrData)
	{
		arrData = arrData || [];
		const nLength = Math.max(arrData.length, this.externalReferences.length);
		const arrPromise = [];
		for (let i = 0; i < nLength; i += 1)
		{
			if (this.isLocalDesktop || (arrData[i] && (!arrData[i]["error"] || this.externalReferences[i].isExternalLink())))
			{
				const oPromiseGetter = new CExternalDataPromiseGetter(this.api, this.getExternalReference(i), arrData[i]);
				arrPromise.push(oPromiseGetter.getPromise());
			}
		}
		this.doUpdate(arrPromise);
	};
	CExternalDataLoader.prototype.doUpdate = function (arrPromise)
	{
		const oThis = this;
		Promise.all(arrPromise).then(function (arrValues)
		{
			oThis.fCallback(arrValues);
		});
	}
	CExternalDataLoader.prototype.getExternalReference = function (nId)
	{
		if (this.externalReferences[nId])
		{
			return this.externalReferences[nId].asc_getPath();
		}
	};

	function CExternalDataPromiseGetter(oApi, sExternalReference, oData)
	{
		this.externalReference = sExternalReference;
		this.data = oData;
		this.api = oApi;
		this.fileUrl = this.getFileUrl();
	}

	CExternalDataPromiseGetter.prototype.isLocalDesktop = function ()
	{
		return window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]();
	}
	CExternalDataPromiseGetter.prototype.resolveStream = function (arrStream, fResolve)
	{
		fResolve({stream: arrStream, externalReferenceId: this.externalReference, data: this.data});
	}
	CExternalDataPromiseGetter.prototype.getLocalDesktopPromise = function ()
	{
		const oThis = this;
		return new Promise(function (resolve)
		{
			if (this.fileUrl)
			{
				window["AscDesktopEditor"]["convertFile"](this.fileUrl, 0x2002, function (_file)
				{
					let arrStream = null;
					if (_file)
					{
						arrStream = _file["get"]();
						_file["close"]();
					}
					oThis.resolveStream(arrStream, resolve);
				});
			}
		});
	};

	CExternalDataPromiseGetter.prototype.getLocalFileLink = function ()
	{
		let res = this.fileUrl;
		if (res)
		{
			res = res.replace(/^file:\/\/\//, '');
			res = res.replace(/^file:\/\//, '');
		}
		return res;
	};

	CExternalDataPromiseGetter.prototype.isExternalLink = function ()
	{
		const p = /^(?:http:|https:)/;
		return this.fileUrl.match(p);
	};

	CExternalDataPromiseGetter.prototype.getFileUrl = function ()
	{
		if (this.isLocalDesktop() && !this.isExternalLink())
		{
			return this.getLocalFileLink();
		}
		else if (this.data && !this.data["error"])
		{
			return this.data["url"];
		}
		return this.externalReference;
	};
	CExternalDataPromiseGetter.prototype.isXlsx = function ()
	{
		const p = /^.*\.(xlsx)$/i;
		return this.fileUrl.match(p);
	};
	CExternalDataPromiseGetter.prototype.isSupportOOXML = function ()
	{
		return this.api["asc_isSupportFeature"]("ooxml");
	};
	CExternalDataPromiseGetter.prototype.getPromise = function ()
	{
		if (this.isLocalDesktop())
		{
			return this.getLocalDesktopPromise();
		}
		else if (!window["NATIVE_EDITOR_ENJINE"])
		{
			return this.getSDKPromise();
		}
	};

	CExternalDataPromiseGetter.prototype.loadFileContentFromUrl = function (sFileUrl, resolve)
	{
		const oThis = this;
		AscCommon.loadFileContent(sFileUrl, function (httpRequest)
		{
			let arrStream = null;
			if (httpRequest)
			{
				arrStream = AscCommon.initStreamFromResponse(httpRequest);
			}
			oThis.resolveStream(arrStream, resolve);
		}, "arraybuffer");
	};

	CExternalDataPromiseGetter.prototype.getSDKPromise = function ()
	{
		const oThis = this;
		return new Promise(function (fResolve)
		{
			const bIsXLSX = oThis.isXlsx();
			const nOutputFormat = oThis.isSupportOOXML() ? Asc.c_oAscFileType.XLSX : Asc.c_oAscFileType.XLSY;
			const sFileUrl = oThis.getFileUrl();
			const sFileType = oThis.data["fileType"];
			const sToken = oThis.data["token"];
			const sDirectUrl = oThis.data["directUrl"];

			if ((sFileUrl && !bIsXLSX) || !oThis.isSupportOOXML())
			{
				let bLoad = false;
				oThis.api.getConvertedXLSXFileFromUrl(sFileUrl, sFileType, sToken, nOutputFormat,
					function (sFileUrlAfterConvert)
					{
						if (sFileUrlAfterConvert)
						{
							oThis.loadFileContentFromUrl(sFileUrlAfterConvert, fResolve);
							bLoad = true;
						}
						else if (!bLoad)
						{
							oThis.resolveStream(null);
						}
					});
			}
			else if (sDirectUrl || sFileUrl)
			{
				oThis.api._downloadOriginalFile(sDirectUrl, sFileUrl, sFileType, sToken, function (arrStream)
				{
					oThis.resolveStream(arrStream, fResolve);
				});
			}
			else
			{
				oThis.resolveStream(null, fResolve);
			}
		});
	};

	function CDiagramUpdater(oApi, oChart)
	{
		this.chart = oChart;
		this.api = oApi;
		this.frameLoader = new CFrameDiagramBinaryLoader(this.api, this.chart);
	}

	CDiagramUpdater.prototype.update = function ()
	{
		const oPromise = this.frameLoader.getPromise();
		if (oPromise)
		{
			oPromise.then(this.resolvePromise.bind(this));
		}
	};

	CDiagramUpdater.prototype.getChartBinary = function ()
	{
		return new Asc.asc_CChartBinary(this.chart);
	};

	CDiagramUpdater.prototype.resolvePromise = function (sStream)
	{
		if (sStream)
		{
			const oBinaryData = this.getChartBinary();
			oBinaryData.setWorkbookBinary(sStream);
			this.sendToFrameEditor(oBinaryData);
		}
		else
		{
			// todo
		}
	};

	CDiagramUpdater.prototype.sendToFrameEditor = function (oBinary)
	{
		this.api.sendFromGeneralToFrameEditor(new CGeneralUpdateDiagramData(oBinary));
	}


	AscCommon.CDiagramCellFrameManager = CDiagramCellFrameManager;
	AscCommon.COleCellFrameManager = COleCellFrameManager;
	AscCommon.CFrameDiagramBinaryLoader = CFrameDiagramBinaryLoader;
	AscCommon.CDiagramUpdater = CDiagramUpdater;
	AscCommon.CFrameUpdateDiagramData = CFrameUpdateDiagramData;
})(window);
