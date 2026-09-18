//! WHATWG Web-platform import signatures.
//!
//! These are the binary-side mirror of `platforms/web`: the runtime handlers
//! still own behavior, while this table tells the Wasm writer which calls are
//! value-producing queries and which are effect-only methods.

use crate::encoding::*;

pub const CANVAS_MODULE: &str = "web:canvas";
pub const DOM_MODULE: &str = "web:dom";
pub const HTML_MODULE: &str = "web:html";
pub const CSSOM_MODULE: &str = "web:cssom";

fn sig(out: &mut Vec<u8>, params: u32, result: Option<u8>) {
    write_leb128_u32(out, params);
    for _ in 0..params {
        out.push(TYPE_EXTERNREF);
    }
    match result {
        Some(ty) => {
            write_leb128_u32(out, 1);
            out.push(ty);
        }
        None => write_leb128_u32(out, 0),
    }
}

pub fn write_signature(out: &mut Vec<u8>, module: &str, name: &str) -> bool {
    match module {
        CANVAS_MODULE => write_canvas_signature(out, name),
        DOM_MODULE => write_dom_signature(out, name),
        HTML_MODULE => write_html_signature(out, name),
        CSSOM_MODULE => write_cssom_signature(out, name),
        _ => false,
    }
}

fn write_canvas_signature(out: &mut Vec<u8>, name: &str) -> bool {
    let (params, result) = match name {
        // Constructors / queries.
        "getContext" => (2, Some(TYPE_EXTERNREF)),
        "createImageData" => (3, Some(TYPE_EXTERNREF)),
        "measureText" => (2, Some(TYPE_EXTERNREF)),
        "getLineDash" | "getTransform" | "getContextAttributes" => (1, Some(TYPE_EXTERNREF)),
        "getImageData" => (5, Some(TYPE_EXTERNREF)),
        "toDataURL" | "toBlob" => (3, Some(TYPE_EXTERNREF)),
        "createLinearGradient" => (5, Some(TYPE_EXTERNREF)),
        "createRadialGradient" => (7, Some(TYPE_EXTERNREF)),
        "createConicGradient" => (4, Some(TYPE_EXTERNREF)),
        "createPattern" => (3, Some(TYPE_EXTERNREF)),
        "createPath2D" => (1, Some(TYPE_EXTERNREF)),

        // Boolean queries.
        "isPointInPath" | "isPointInStroke" => (4, Some(TYPE_I32)),
        "isContextLost" => (1, Some(TYPE_I32)),

        // String-valued reads.
        "getFont"
        | "getFillStyle"
        | "getStrokeStyle"
        | "getFilter"
        | "getGlobalCompositeOperation"
        | "getImageSmoothingQuality"
        | "getShadowColor"
        | "getDirection"
        | "getLetterSpacing"
        | "getWordSpacing"
        | "getFontKerning"
        | "getFontStretch"
        | "getFontVariantCaps"
        | "getTextRendering"
        | "getLang"
        | "getTextAlign"
        | "getTextBaseline"
        | "getLineCap"
        | "getLineJoin" => (1, Some(TYPE_EXTERNREF)),

        // Effect-only drawing, state, path, gradient mutation and blit calls.
        "save" | "restore" | "resetTransform" | "beginPath" | "closePath" | "reset"
        | "clearAll" | "pathClosePath" => (1, None),
        "setLineWidth"
        | "setLineCap"
        | "setLineJoin"
        | "setGlobalAlpha"
        | "setImageSmoothingEnabled"
        | "rotate"
        | "setMiterLimit"
        | "setTextAlign"
        | "setTextBaseline"
        | "setLineDashOffset"
        | "setFillStyleCss"
        | "setStrokeStyleCss"
        | "setFontCss"
        | "setFilter"
        | "setShadowColor"
        | "setGlobalCompositeOperation"
        | "fillWithRule"
        | "clipWithRule"
        | "fill"
        | "clip"
        | "stroke"
        | "appendPath"
        | "drawFocusIfNeeded"
        | "addColorStop"
        | "addPath" => (2, None),
        "setFillStyle" | "setStrokeStyle" => (5, None),
        "translate" | "scale" | "moveTo" | "lineTo" | "pathMoveTo" | "pathLineTo" => (3, None),
        "setFont" => (5, None),
        "transform" | "setTransform" | "bezierCurveTo" | "pathBezierCurveTo" => (7, None),
        "ellipse" | "pathEllipse" => (9, None),
        "setLineDash" => (8, None),
        "arc" | "pathArc" => (7, None),
        "quadraticCurveTo" | "rect" | "fillRect" | "strokeRect" | "clearRect"
        | "fillTextMaxWidth" | "strokeTextMaxWidth" | "pathQuadraticCurveTo" | "pathRect"
        | "pathArcTo" | "pathRoundRect" => (5, None),
        "fillText" | "strokeText" | "putImageData" => (4, None),
        "drawImage" => (6, None),
        _ => return false,
    };
    sig(out, params, result);
    true
}

fn write_dom_signature(out: &mut Vec<u8>, name: &str) -> bool {
    let (params, result) = match name {
        "createElement" | "createElementNS" | "createTextNode" | "createComment" => {
            (3, Some(TYPE_EXTERNREF))
        }
        "appendChild" | "removeChild" => (3, Some(TYPE_EXTERNREF)),
        "getElementById" | "querySelector" | "querySelectorAll" | "getElementsByTagName" => {
            (2, Some(TYPE_EXTERNREF))
        }
        "textContent" | "innerHtml" | "getAttributeNames" => (2, Some(TYPE_EXTERNREF)),
        "getAttribute" => (3, Some(TYPE_EXTERNREF)),
        "isConnected" => (2, Some(TYPE_I32)),
        "hasAttribute" | "matches" | "contains" => (3, Some(TYPE_I32)),
        "closest"
        | "firstChild"
        | "lastChild"
        | "nextSibling"
        | "previousSibling"
        | "parentNode"
        | "childNodes"
        | "children"
        | "firstElementChild"
        | "lastElementChild" => (1, Some(TYPE_EXTERNREF)),
        "setTextContent" | "setInnerHtml" => (3, None),
        "removeAttribute" => (3, None),
        "toggleAttribute" => (4, None),
        "setAttribute" | "addEventListener" | "removeEventListener" | "classListAdd"
        | "classListRemove" | "classListToggle" => (4, None),
        "setAttributeNS" => (4, None),
        "classListContains" => (3, Some(TYPE_I32)),
        "remove" => (1, None),
        _ => return false,
    };
    sig(out, params, result);
    true
}

fn write_html_signature(out: &mut Vec<u8>, name: &str) -> bool {
    let (params, result) = match name {
        "activeDocument" => (0, Some(TYPE_EXTERNREF)),
        "body" | "title" | "value" | "checked" | "defaultView" => (1, Some(TYPE_EXTERNREF)),
        "focus" | "showPicker" | "show" | "showModal" | "close" => (1, None),
        "setTitle" => (2, None),
        _ => return false,
    };
    sig(out, params, result);
    true
}

fn write_cssom_signature(out: &mut Vec<u8>, name: &str) -> bool {
    let (params, result) = match name {
        "getPropertyValue" => (2, Some(TYPE_EXTERNREF)),
        "removeProperty" => (2, None),
        "setProperty" => (3, None),
        _ => return false,
    };
    sig(out, params, result);
    true
}
