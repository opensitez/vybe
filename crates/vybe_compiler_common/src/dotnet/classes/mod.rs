//! .NET-shaped class wrappers for the GUI surface.
//!
//! This submodule defines the metadata + chunk-building helpers that turn
//! the .NET WinForms class hierarchy into real compile-time classes
//! registered against the user's compiler. Each class becomes a real
//! constructor chunk with a real parent-call chain (via the existing
//! `compile_class` machinery), so user code that writes
//!
//! ```vb
//! Public Class Form1
//!     Inherits Form
//!     Public Sub New()
//!         Me.Text = "Hello"
//!     End Sub
//! End Class
//! ```
//!
//! gets a real `Form` parent class with a real ctor that walks the .NET
//! inheritance chain (`Form → ContainerControl → ScrollableControl →
//! Control → Component → MarshalByRefObject → Object`) and binds the
//! property setters owned at each level. The user write `Me.Text = "Hello"`
//! emits plain `struct_set "text"`; the VM finds `__set_text` (installed
//! by `Control` in the inherited chain) and dispatches to a setter chunk
//! that calls `vybe:gui::controlSetProperty(this, "Text", "Hello")`.
//!
//! ## Why classes and not host-side setter installation
//!
//! The previous shortcut was to install `__set_<prop>` setter closures on
//! every control object inside the `vybe:gui::new_<Type>` host fn. That
//! "worked" but flattened the inheritance, hardcoded the property surface
//! into the host, and gave user code no real `Control`/`Form`/etc. class
//! identity to inherit from. Real classes give us:
//!
//! - `obj is Control`, `obj is Form`, `MyBase` / `base.X()` calls
//! - User subclassing of any level (`class MyButton : Inherits Button`)
//! - Properties placed at the .NET-correct level (Text on Control,
//!   FormBorderStyle on Form, DialogResult on Button)
//! - The inheritance chain is the single source of truth — if .NET adds a
//!   property, we add it once at the right level
//!
//! ## Module layout
//!
//! ```text
//! dotnet/classes/
//! ├── mod.rs          — DotnetClass struct, registration entry, table
//! ├── builder.rs      — chunk-building helpers (private)
//! ├── object.rs       — Object, MarshalByRefObject, Component
//! ├── control.rs      — Control, ScrollableControl, ContainerControl
//! ├── form.rs         — Form
//! ├── buttons.rs      — ButtonBase, Button, CheckBox, RadioButton
//! ├── text.rs         — TextBoxBase, TextBox, RichTextBox, MaskedTextBox
//! ├── labels.rs       — Label, LinkLabel
//! ├── lists.rs        — ListControl, ComboBox, ListBox, ListView, TreeView
//! ├── containers.rs   — Panel, GroupBox, TabControl, TabPage, …
//! ├── progress.rs     — ProgressBar, TrackBar, NumericUpDown
//! ├── dates.rs        — DateTimePicker, MonthCalendar
//! ├── media.rs        — PictureBox, WebBrowser
//! ├── grids.rs        — DataGridView
//! ├── strips.rs       — MenuStrip, ToolStrip, StatusStrip, ContextMenuStrip
//! ├── nonvisual.rs    — Timer, BindingSource, ImageList, ToolTip, …
//! └── dialogs.rs      — OpenFileDialog, SaveFileDialog, …
//! ```
//!
//! Family files are added phase-by-phase. Each one returns a slice of
//! `DotnetClass` definitions; `dotnet_classes()` collects them all.

pub mod builder;
pub mod object;
pub mod control;
pub mod form;
pub mod buttons;
pub mod text;
pub mod labels;
pub mod lists;
pub mod containers;
pub mod progress;
pub mod dates;
pub mod media;
pub mod grids;
pub mod strips;
pub mod nonvisual;
pub mod dialogs;

/// Metadata for a single .NET BCL class wrapper.
///
/// The compiler walks the class table, topologically sorts by inheritance,
/// and generates one constructor chunk per class. Each constructor chunk
/// calls its parent's ctor (via the existing parent-call machinery), then
/// binds one setter chunk per property in `properties`. Concrete leaves
/// (those with `widget_host_fn = Some(_)`) additionally call the matching
/// `vybe:gui::new_<Type>` host fn at the bottom of their ctor and copy the
/// resulting widget identity (`__control_name`, `__control_type`, `name`,
/// and the `show`/`close`/`focus`/`hide` method refs) onto `this`.
#[derive(Debug, Clone, Copy)]
pub struct DotnetClass {
    /// Canonical .NET name (PascalCase). Used as the global name and as
    /// the value stamped into `__type` by this class's ctor.
    pub name: &'static str,

    /// Parent class name. `None` only for `Object`.
    pub parent: Option<&'static str>,

    /// Properties added at this class level (NOT inherited). The setter
    /// for each gets bound under `__set_<lowercased_name>` and calls
    /// `vybe:gui::controlSetProperty(this, "<name>", value)` so the value
    /// mirrors into the host gui state registry.
    ///
    /// Names are written in PascalCase (`"Text"`, `"FormBorderStyle"`)
    /// because that's what `controlSetProperty` receives as the property
    /// key — the .NET-canonical form. The setter binding uses the
    /// lowercased form to match the VM's `struct_set → __set_<field>`
    /// dispatch (struct fields are lowercased on the canonical AST path).
    pub properties: &'static [&'static str],

    /// If `Some(host_fn)`, this is a concrete leaf class — its ctor calls
    /// `vybe:gui::<host_fn>` at the end (after the parent chain has bound
    /// all the inherited setters) to wire the actual vybe_widgets backing
    /// widget. The widget's identity (`__control_name`, `__control_type`,
    /// `name`, `show`/`close`/`focus`/`hide` method refs) is copied onto
    /// `this`, leaving the inherited setters intact.
    ///
    /// Abstract bases (`Object`, `Control`, `ButtonBase`, `TextBoxBase`,
    /// `ListControl`, `ScrollableControl`, `ContainerControl`, …) leave
    /// this `None` — they bind setters but don't materialize a widget.
    pub widget_host_fn: Option<&'static str>,
}

impl DotnetClass {
    /// True if this class has an underlying vybe_widgets host backing.
    pub fn is_concrete(&self) -> bool {
        self.widget_host_fn.is_some()
    }
}

/// The complete .NET class table.
///
/// Each family file (`object`, `control`, `form`, `buttons`, …) contributes
/// a slice via its `classes()` function. They are concatenated here in
/// inheritance-friendly order — no topological sort needed because each
/// family is internally ordered (`Object` before `MarshalByRefObject`
/// before `Component`, …) and parents are listed in earlier families
/// than their children.
pub fn dotnet_classes() -> &'static [DotnetClass] {
    use std::sync::LazyLock;
    static TABLE: LazyLock<Vec<DotnetClass>> = LazyLock::new(|| {
        let mut v = Vec::new();
        v.extend_from_slice(object::classes());
        v.extend_from_slice(control::classes());
        v.extend_from_slice(form::classes());
        v.extend_from_slice(buttons::classes());
        v.extend_from_slice(text::classes());
        v.extend_from_slice(labels::classes());
        v.extend_from_slice(lists::classes());
        v.extend_from_slice(containers::classes());
        v.extend_from_slice(progress::classes());
        v.extend_from_slice(dates::classes());
        v.extend_from_slice(media::classes());
        v.extend_from_slice(grids::classes());
        v.extend_from_slice(strips::classes());
        v.extend_from_slice(nonvisual::classes());
        v.extend_from_slice(dialogs::classes());
        v
    });
    &TABLE
}
