//! Abstract base classes — the `is`/type-identity chain. No backing controls
//! and no fields; they exist so `x is StatelessWidget` / `is ScrollView` etc.
//! resolve by ancestry membership.

use crate::emitter::catalog::FlutterClass;

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::abstract_("Widget", None),
    FlutterClass::abstract_("StatelessWidget", Some("Widget")),
    FlutterClass::abstract_("StatefulWidget", Some("Widget")),
    FlutterClass::abstract_("RenderObjectWidget", Some("Widget")),
    FlutterClass::abstract_("SingleChildRenderObjectWidget", Some("RenderObjectWidget")),
    FlutterClass::abstract_("MultiChildRenderObjectWidget", Some("RenderObjectWidget")),
    FlutterClass::abstract_("ProxyWidget", Some("Widget")),
    FlutterClass::abstract_("ParentDataWidget", Some("ProxyWidget")),
    FlutterClass::abstract_("InheritedWidget", Some("ProxyWidget")),
    FlutterClass::abstract_("PreferredSizeWidget", Some("Widget")),
    FlutterClass::abstract_("ScrollView", Some("StatelessWidget")),
    FlutterClass::abstract_("BoxScrollView", Some("ScrollView")),
    FlutterClass::abstract_("FormField", Some("StatefulWidget")),
    FlutterClass::abstract_("ImplicitlyAnimatedWidget", Some("StatefulWidget")),
    FlutterClass::abstract_("AnimatedWidget", Some("StatefulWidget")),
];
