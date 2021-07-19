//	@ghasemkiani/word-maker/word

const {cutil} = require("@ghasemkiani/base/cutil");
const {Page} = require("@ghasemkiani/htmlmaker/page");
const {WDocument} = require("@ghasemkiani/wjsdom/document");

class Word extends Page {
	render({wnode}) {
		let res = super.render({wnode});
		return {...res};
	}
	renderHtml(arg) {
		arg = Object.assign({
			wnode: new WDocument().root,
			cs: "UTF-8",
			lang: "en-US",
			tab: "24pt",
			title: null,
		}, arg);
		let {wnode} = arg;
		let wnodeHtml;
		let wnodeHead;
		let wnodeBody;
		wnode.chain(wnode => {
			wnodeHtml = wnode;
			wnode.attr({
				"xmlns:v": "urn:schemas-microsoft-com:vml",
				"xmlns:o": "urn:schemas-microsoft-com:office:office",
				"xmlns:w": "urn:schemas-microsoft-com:office:this",
				"xmlns:m": "http://schemas.microsoft.com/office/2004/12/omml",
			});
			wnode.ch("head", wnode => {
				wnodeHead = wnode;
				wnode.ch("meta[http-equiv=Content-Type]", wnode => {
					wnode.attr("content", `text/html;charset=${arg.cs}`);
				});
				wnode.ch("meta[name=ProgId,content=Word.Document]");
				wnode.c("xml", "", wnode => {
					wnode.css("display", "none");
					wnode.c("w:WordDocument", wnode => {
						wnode.c("w:View$Print");
						wnode.c("w:Zoom$BestFit");
					});
				});
				if (!cutil.isNil(arg.title)) {
					wnode.ch("title", wnode => {
						wnode.t(arg.title);
					});
				}
			});
			wnode.ch("body", wnode => {
				wnodeBody = wnode;
				wnode.attr("lang", arg.lang);
				wnode.css("tab-interval", arg.tab);
			});
		});
		return {wnode, wnodeHtml, wnodeHead, wnodeBody};
	}
	renderPageBreak({wnode}) {
		wnode.chain(wnode => {
			wnode.ch("br[clear=all]{mso-special-character:line-break;page-break-before:always}");
		});
		return {wnode};
	}
	renderField({wnode, ...arg}) {
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
	renderFieldPageRef({wnode, ...arg}) {
		arg = juya.require("gk/type").construct({
				bookmark: null,
				dontLink: false,
				onResult(wnode) {
					wnode.t("*");
				},
			}).assign(arg);
		let {bookmark, dontLink} = arg;
		return this.renderField({
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
	renderFieldRef({wnode, ...arg}) {
		arg = juya.require("gk/type").construct({
				bookmark: null,
				dontLink: false,
			}).assign(arg);
		let {bookmark, dontLink} = arg;
		return this.renderField({
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
	renderFieldSet({wnode, ...arg}) {
		arg = juya.require("gk/type").construct({
				bookmark: null,
				onValue(wnode) {},
			}).assign(arg);
		let {bookmark, onValue} = arg;
		return this.renderField({
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
	renderFieldIf({wnode, ...arg}) {
		arg = juya.require("gk/type").construct({
				bookmark: null,
				onCondition(wnode) {},
				onValue1(wnode) {},
				onValue2(wnode) {},
			}).assign(arg);
		let {bookmark, onValue} = arg;
		return this.renderField({
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
	renderPageRefList({wnode, ...arg}) {
		let word = this;
		arg = Object.assign({
				refs: [],
				delimiter: "\u060C ",
			}, arg);
		let {refs, delimiter} = arg;
		let name = refs[0];
		this.renderFieldSet({
			wnode,
			bookmark: "idxpg",
			onValue(wnode) {
				wnode.chain(function (wnode) {
					word.renderFieldPageRef({
						wnode,
						bookmark: name,
					});
				});
			},
		});
		this.renderFieldRef({
			wnode,
			bookmark: "idxpg",
		});
		for(let name of refs.slice(1)) {
			this.makeFieldSet({
				wnode,
				bookmark: "idxpg1",
				onValue(wnode) {
					word.renderFieldPageRef({
						wnode,
						bookmark: name,
					});
				},
			});
			this.renderFieldIf({
				wnode,
				onCondition(wnode) {
					wnode.t("idxpg = idxpg1");
				},
				onValue1(wnode) {},
				onValue2(wnode) {
					wnode.ch("span[dir=rtl]", wnode => {
						wnode.t(delimiter);
					});
					word.renderFieldSet({
						wnode,
						bookmark: "idxpg",
						onValue(arg) {
							word.renderFieldRef({
								wnode,
								bookmark: "idxpg1",
							});
						},
					});
					word.renderFieldRef({
						wnode,
						bookmark: "idxpg",
					});
				},
			});
		}
		return word;
	}
	renderImg({wnode, ...arg}) {
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
	renderTab({wnode}) {
		let wnodeSpan;
		wnode.ch("span[style=mso-tab-count:1;]", wnode => {
			wnodeSpan = wnode;
			wnode.t("\t");
		});
		return {wnode, wnodeSpan};
	}
}
cutil.extend(Word.prototype, {
	//
});

module.exports = {Word};
