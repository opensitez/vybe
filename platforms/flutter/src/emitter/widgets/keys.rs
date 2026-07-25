//! Key hierarchy — `Key` and its `LocalKey`/`GlobalKey` subtypes. Value-typed
//! (structural `==`); `Key('x')` is a factory for `ValueKey<String>('x')`.

use crate::emitter::catalog::{FlutterClass, FlutterField};

const F_VALKEY: &[FlutterField] = &[FlutterField::positional("value", 0)];
const F_GLOBALKEY: &[FlutterField] = &[FlutterField::named("debugLabel")];

pub(crate) const CLASSES: &[FlutterClass] = &[
    // `Key('x')` constructs a value key carrying ValueKey/LocalKey identity so
    // `is ValueKey`/`is LocalKey` match; value equality via the allow-list.
    FlutterClass::data_with_interfaces("Key", None, &["ValueKey", "LocalKey"], F_VALKEY),
    FlutterClass::abstract_("LocalKey", Some("Key")),
    FlutterClass::data("ValueKey", Some("LocalKey"), F_VALKEY),
    FlutterClass::data("ObjectKey", Some("LocalKey"), F_VALKEY),
    FlutterClass::data("UniqueKey", Some("LocalKey"), &[]),
    FlutterClass::data("GlobalKey", Some("Key"), F_GLOBALKEY),
    FlutterClass::data("GlobalObjectKey", Some("GlobalKey"), F_VALKEY),
];
