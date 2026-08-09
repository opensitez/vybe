//! Image widget & image providers — the `Image` widget lowers onto a
//! `vybe:gui` picture box; the providers are pure data captured at construction.

use crate::emitter::catalog::{FlutterClass, FlutterField};

const F_IMAGE: &[FlutterField] = &[
    FlutterField::named("image"),
    FlutterField::named("width"),
    FlutterField::named("height"),
    FlutterField::named("fit"),
    FlutterField::named("alignment"),
    FlutterField::named("repeat"),
    FlutterField::named("color"),
    FlutterField::named("colorBlendMode"),
    FlutterField::named("filterQuality"),
];

const F_NETIMAGE: &[FlutterField] = &[
    FlutterField::positional("url", 0),
    FlutterField::named("scale"),
];

const F_ASSETIMAGE: &[FlutterField] = &[
    FlutterField::positional("assetName", 0),
    FlutterField::named("scale"),
];

const F_FILEIMAGE: &[FlutterField] = &[FlutterField::positional("file", 0)];

const F_MEMIMAGE: &[FlutterField] = &[FlutterField::positional("bytes", 0)];

const F_RESIZEIMAGE: &[FlutterField] = &[
    FlutterField::positional("imageProvider", 0),
    FlutterField::named("width"),
    FlutterField::named("height"),
];

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::widget("Image", "StatefulWidget", "picturebox", F_IMAGE),
    FlutterClass::data("NetworkImage", None, F_NETIMAGE),
    FlutterClass::data("AssetImage", None, F_ASSETIMAGE),
    FlutterClass::data("ExactAssetImage", None, F_ASSETIMAGE),
    FlutterClass::data("FileImage", None, F_FILEIMAGE),
    FlutterClass::data("MemoryImage", None, F_MEMIMAGE),
    FlutterClass::data("ResizeImage", None, F_RESIZEIMAGE),
];
