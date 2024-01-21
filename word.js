import {cutil} from "@ghasemkiani/base";
import {Maker} from "@ghasemkiani/html-maker";
import {iwxhtml} from "@ghasemkiani/html-maker";
import {iwjsdom} from "@ghasemkiani/jsdom";

const XMLNS_V = "urn:schemas-microsoft-com:vml";
const XMLNS_O = "urn:schemas-microsoft-com:office:office";
const XMLNS_W = "urn:schemas-microsoft-com:office:this";
const XMLNS_M = "http://schemas.microsoft.com/office/2004/12/omml";

class Word extends cutil.mixin(Maker, iwxhtml, iwjsdom) {
	static {
		cutil.extend(this, {
			XMLNS_V,
			XMLNS_O,
			XMLNS_W,
			XMLNS_M,
		});
		cutil.extend(this.prototype, {
			XMLNS_V,
			XMLNS_O,
			XMLNS_W,
			XMLNS_M,
		});
	}
	makeDoc({cs = "UTF-8", lang = "en-US", tab = "24pt", ...arg}) {
		let maker = this;
		let {x} = maker;
		let {xhtml} = maker;
		let {document, nhtml, nhead, ntitle, nbody, ndescription, nkeywords, nauthor} = xhtml.makeDoc({noMetaCharset: true, ...cutil.asObject(arg)});
		
		x.chain(nhtml, node => {
			if (cutil.na(x.attr(node, "xmlns:v"))) {
				x.attr(node, "xmlns:v", XMLNS_V);
			}
			if (cutil.na(x.attr(node, "xmlns:o"))) {
				x.attr(node, "xmlns:o", XMLNS_O);
			}
			if (cutil.na(x.attr(node, "xmlns:w"))) {
				x.attr(node, "xmlns:w", XMLNS_W);
			}
			if (cutil.na(x.attr(node, "xmlns:m"))) {
				x.attr(node, "xmlns:m", XMLNS_M);
			}
			x.chain(nhead, node => {
				x.cx(node, "meta[http-equiv=Content-Type]", node => {
					x.attr(node, "content", `text/html;charset=${cs}`);
				});
				x.cx(node, "meta[name=ProgId,content=Word.Document]");
				x.c(node, "xml", node => {
					x.css(node, {"display": "none"});
					x.cx(node, "w:WordDocument", XMLNS_W, node => {
						x.cx(node, "w:View$Print", XMLNS_W);
						x.cx(node, "w:Zoom$BestFit", XMLNS_W);
					});
				});
			});
			x.chain(nbody, node => {
				x.attr(node, "lang", lang);
				x.css(node, {"tab-interval": tab});
			});
		});
		return {nhtml, nhead, nbody};
	}
	makePageBreak({node}) {
		let maker = this;
		let {x} = maker;
		let nbr;
		x.chain(node, node => {
			nbr = x.cx(node, "br[clear=all]{mso-special-character:line-break;page-break-before:always;}");
		});
		return {node, nbr};
	}
	makeField({
		node,
		onField,
		onResult,
	}) {
		let maker = this;
		let {x} = maker;
		onField ||= node => {
			x.t(node, "Quote");
		};
		onResult ||= node => {
			x.t(node, "*");
		};
		let nfieldBegin;
		let nfieldCode;
		let nfieldSeparator;
		let nfieldEnd;
		x.chain(node, node => {
			x.cx(node, "span[style=mso-element:field-begin;]", node => {
				nfieldBegin = node;
			});
			x.cx(node, "span[dir=ltr]", node => {
				nfieldCode = node;
				x.chain(node, onField);
			});
			x.cx(node, "span[style=mso-element:field-separator;]", node => {
				nfieldSeparator = node;
			});
			x.chain(node, onResult);
			x.cx(node, "span[style=mso-element:field-end;]", node => {
				nfieldEnd = node;
			});
		});
		return {node, nfieldBegin, nfieldCode, nfieldSeparator, nfieldEnd};
	}
	makeFieldPageRef({
		node,
		bookmark = null,
		dontLink = false,
		onResult,
	}) {
		let maker = this;
		let {x} = maker;
		onResult ||= node => {
			x.t(node, "*");
		};
		return maker.makeField({
			node,
			onField(node) {
				x.cx(node, "span[dir=ltr]$pageref");
				if (!dontLink) {
					x.t(node, " ");
					x.cx(node, "span[dir=ltr]$\\h");
				}
				x.t(node, " ");
				x.cx(node, "span[dir=ltr]", node => {
					x.t(node, bookmark);
				});
			},
			onResult,
		});
	}
	makeFieldRef({
		node,
		bookmark = null,
		dontLink = false,
	}) {
		let maker = this;
		let {x} = maker;
		return maker.makeField({
			node,
			onField(node) {
				x.cx(node, "span[dir=ltr]$ref");
				if (!dontLink) {
					x.t(node, " ");
					x.cx(node, "span[dir=ltr]$\\h");
				}
				x.t(node, " ");
				x.cx(node, "span[dir=ltr]", node => {
					x.t(node, bookmark);
				});
			},
		});
	}
	makeFieldSet({
		node,
		bookmark = null,
		onValue = node => {},
	}) {
		let maker = this;
		let {x} = maker;
		return maker.makeField({
			node,
			onField(node) {
				x.cx(node, "span[dir=ltr]$set");
				x.t(node, " ");
				x.cx(node, "span[dir=ltr]", node => {
					x.t(node, bookmark);
				});
				x.t(node, " ");
				x.t(node, '"');
				x.chain(node, onValue);
				x.t(node, '"');
			},
		});
	}
	makeFieldIf({
		node,
		bookmark = null,
		onCondition = node => {},
		onValue1 = node => {},
		onValue2 = node => {},
	}) {
		let maker = this;
		let {x} = maker;
		return maker.makeField({
			node,
			onField(node) {
				x.cx(node, "span[dir=ltr]$if");
				x.t(node, " ");
				x.cx(node, "span[dir=ltr]", node => {
					x.chain(node, onCondition);
				});
				x.t(node, " ");
				x.t(node, '"');
				x.chain(node, onValue1);
				x.t(node, '"');
				x.t(node, " ");
				x.t(node, '"');
				x.chain(node, onValue2);
				x.t(node, '"');
			},
		});
	}
	makePageRefList({
		node,
		refs = [],
		delimiter = "\u060C ",
	}) {
		let maker = this;
		let {x} = maker;
		let name = refs[0];
		maker.makeFieldSet({
			node,
			bookmark: "idxpg",
			onValue(node) {
				x.chain(node, node => {
					maker.makeFieldPageRef({node, bookmark: name});
				});
			},
		});
		maker.makeFieldRef({node, bookmark: "idxpg"});
		for(let name of refs.slice(1)) {
			maker.makeFieldSet({
				node,
				bookmark: "idxpg1",
				onValue(node) {
					word.makeFieldPageRef({
						node,
						bookmark: name,
					});
				},
			});
			this.makeFieldIf({
				node,
				onCondition(node) {
					x.t(node, "idxpg = idxpg1");
				},
				onValue1(node) {},
				onValue2(node) {
					x.cx(node, "span[dir=rtl]", node => {
						x.t(node, delimiter);
					});
					maker.makeFieldSet({
						node,
						bookmark: "idxpg",
						onValue(arg) {
							maker.makeFieldRef({node, bookmark: "idxpg1"});
						},
					});
					maker.makeFieldRef({node, bookmark: "idxpg"});
				},
			});
		}
		return maker;
	}
	makeImg({
		node,
		url = null,
		dpi = 300,
		targetDpi = 96,
		width = 0,
		height = 0,
	}) {
		let maker = this;
		let {x} = maker;
		let nodeImg;
		x.cx(node, "img", node => {
			nodeImg = node;
			x.attr(node, "src", url);
			x.attr(node, "width", width * targetDpi / dpi);
			x.attr(node, "height", height * targetDpi / dpi);
		});
		return {node, nodeImg};
	}
	static makeTab({node}) {
		let maker = this;
		let {x} = maker;
		let nspan;
		x.cx(node, "span[style=mso-tab-count:1;]", node => {
			nspan = node;
			x.t(node, "\t");
		});
		return {node, nspan};
	}
}

export {Word};
