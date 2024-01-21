import {cutil} from "@ghasemkiani/base";
import {Obj} from "@ghasemkiani/base";
import {Html} from "@ghasemkiani/html-maker";
import {WDocument} from "@ghasemkiani/wjsdom";

class Word extends Obj {
	static {
		cutil.extend(this.prototype, {
			//
		});
	}
	static makeDoc({
			whtml = new WDocument().root,
			cs = "UTF-8",
			lang = "en-US",
			tab = "24pt",
			title = null,
		}) {
		let whead;
		let wbody;
		whtml.chain(wnode => {
			whtml = wnode;
			wnode.attr({
				"xmlns:v": "urn:schemas-microsoft-com:vml",
				"xmlns:o": "urn:schemas-microsoft-com:office:office",
				"xmlns:w": "urn:schemas-microsoft-com:office:this",
				"xmlns:m": "http://schemas.microsoft.com/office/2004/12/omml",
			});
			wnode.ch("head", wnode => {
				whead = wnode;
				wnode.ch("meta[http-equiv=Content-Type]", wnode => {
					wnode.attr("content", `text/html;charset=${cs}`);
				});
				wnode.ch("meta[name=ProgId,content=Word.Document]");
				wnode.c("xml", "", wnode => {
					wnode.css("display", "none");
					wnode.c("w:WordDocument", wnode => {
						wnode.c("w:View$Print");
						wnode.c("w:Zoom$BestFit");
					});
				});
				if (!cutil.isNil(title)) {
					wnode.ch("title", wnode => {
						wnode.t(title);
					});
				}
			});
			wnode.ch("body", wnode => {
				wbody = wnode;
				wnode.attr("lang", lang);
				wnode.css("tab-interval", tab);
			});
		});
		return {whtml, whead, wbody};
	}
	static makePageBreak({wnode}) {
		wnode.chain(wnode => {
			wnode.ch("br[clear=all]{mso-special-character:line-break;page-break-before:always}");
		});
		return {wnode};
	}
	static makeField({wnode, ...arg}) {
		arg = Object.assign({
			onField(wnode) {
				wnode.t("Quote");
			},
			onResult(wnode) {
				wnode.t("*");
			},
		}, arg);
		let wnodeFieldBegin;
		let wnodeFieldCode;
		let wnodeFieldSeparator;
		let wnodeFieldEnd;
		wnode.chain(wnode => {
			wnode.ch("span[style=mso-element:field-begin;]", wnode => {
				wnodeFieldBegin = wnode;
			});
			wnode.ch("span[dir=ltr]", wnode => {
				wnodeFieldCode = wnode;
				wnode.chain(arg.onField);
			});
			wnode.ch("span[style=mso-element:field-separator;]", wnode => {
				wnodeFieldSeparator = wnode;
			});
			wnode.chain(arg.onResult);
			wnode.ch("span[style=mso-element:field-end;]", wnode => {
				wnodeFieldEnd = wnode;
			});
		});
		return {wnode, wnodeFieldBegin, wnodeFieldCode, wnodeFieldSeparator, wnodeFieldEnd};
	}
	static makeFieldPageRef({wnode, ...arg}) {
		arg = juya.require("gk/type").construct({
				bookmark: null,
				dontLink: false,
				onResult(wnode) {
					wnode.t("*");
				},
			}).assign(arg);
		let {bookmark, dontLink} = arg;
		return this.makeField({
			wnode,
			onField(wnode) {
				wnode.ch("span[dir=ltr]$pageref");
				if (!dontLink) {
					wnode.t(" ");
					wnode.ch("span[dir=ltr]$\\h");
				}
				wnode.t(" ");
				wnode.ch("span[dir=ltr]", wnode => {
					wnode.t(bookmark);
				});
			},
			onResult: arg.onResult,
		});
	}
	static makeFieldRef({wnode, ...arg}) {
		arg = juya.require("gk/type").construct({
				bookmark: null,
				dontLink: false,
			}).assign(arg);
		let {bookmark, dontLink} = arg;
		return this.makeField({
			wnode,
			onField(wnode) {
				wnode.ch("span[dir=ltr]$ref");
				if (!dontLink) {
					wnode.t(" ");
					wnode.ch("span[dir=ltr]$\\h");
				}
				wnode.t(" ");
				wnode.ch("span[dir=ltr]", wnode => {
					wnode.t(bookmark);
				});
			},
		});
	}
	static makeFieldSet({wnode, ...arg}) {
		arg = juya.require("gk/type").construct({
				bookmark: null,
				onValue(wnode) {},
			}).assign(arg);
		let {bookmark, onValue} = arg;
		return this.makeField({
			wnode,
			onField(wnode) {
				wnode.ch("span[dir=ltr]$set");
				wnode.t(" ");
				wnode.ch("span[dir=ltr]", wnode => {
					wnode.t(bookmark);
				});
				wnode.t(" ");
				wnode.t('"');
				wnode.chain(onValue);
				wnode.t('"');
			},
		});
	}
	static makeFieldIf({wnode, ...arg}) {
		arg = juya.require("gk/type").construct({
				bookmark: null,
				onCondition(wnode) {},
				onValue1(wnode) {},
				onValue2(wnode) {},
			}).assign(arg);
		let {bookmark, onValue} = arg;
		return this.makeField({
			wnode,
			onField(wnode) {
				wnode.ch("span[dir=ltr]$if");
				wnode.t(" ");
				wnode.ch("span[dir=ltr]", wnode => {
					wnode.chain(onCondition);
				});
				wnode.t(" ");
				wnode.t('"');
				wnode.chain(onValue1);
				wnode.t('"');
				wnode.t(" ");
				wnode.t('"');
				wnode.chain(onValue2);
				wnode.t('"');
			},
		});
	}
	static makePageRefList({wnode, ...arg}) {
		let word = this;
		arg = Object.assign({
				refs: [],
				delimiter: "\u060C ",
			}, arg);
		let {refs, delimiter} = arg;
		let name = refs[0];
		this.makeFieldSet({
			wnode,
			bookmark: "idxpg",
			onValue(wnode) {
				wnode.chain(function (wnode) {
					word.makeFieldPageRef({
						wnode,
						bookmark: name,
					});
				});
			},
		});
		this.makeFieldRef({
			wnode,
			bookmark: "idxpg",
		});
		for(let name of refs.slice(1)) {
			this.makeFieldSet({
				wnode,
				bookmark: "idxpg1",
				onValue(wnode) {
					word.makeFieldPageRef({
						wnode,
						bookmark: name,
					});
				},
			});
			this.makeFieldIf({
				wnode,
				onCondition(wnode) {
					wnode.t("idxpg = idxpg1");
				},
				onValue1(wnode) {},
				onValue2(wnode) {
					wnode.ch("span[dir=rtl]", wnode => {
						wnode.t(delimiter);
					});
					word.makeFieldSet({
						wnode,
						bookmark: "idxpg",
						onValue(arg) {
							word.makeFieldRef({
								wnode,
								bookmark: "idxpg1",
							});
						},
					});
					word.makeFieldRef({
						wnode,
						bookmark: "idxpg",
					});
				},
			});
		}
		return word;
	}
	static makeImg({wnode, ...arg}) {
		arg = Object.assign({
				url: null,
				dpi: 300,
				targetDpi: 96,
				width: 0,
				height: 0,
			}, arg);
		let wnodeImg;
		wnode.ch("img", wnode => {
			wnodeImg = wnode;
			wnode.attr("src", arg.url);
			wnode.attr("width", arg.width * arg.targetDpi / arg.dpi);
			wnode.attr("height", arg.height * arg.targetDpi / arg.dpi);
		});
		return {wnode, wnodeImg};
	}
	static makeTab({wnode}) {
		let wnodeSpan;
		wnode.ch("span[style=mso-tab-count:1;]", wnode => {
			wnodeSpan = wnode;
			wnode.t("\t");
		});
		return {wnode, wnodeSpan};
	}
}

export {Word};
