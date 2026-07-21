//! WinForms-shaped class wrappers for the GUI surface.
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
//! dotnet/winforms/classes/
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
pub mod buttons;
pub mod containers;
pub mod control;
pub mod dates;
pub mod dialogs;
pub mod drawing;
pub mod form;
pub mod grids;
pub mod labels;
pub mod lists;
pub mod media;
pub mod nonvisual;
pub mod object;
pub mod progress;
pub mod strips;
pub mod text;

/// Metadata for a single .NET BCL class wrapper.
///
/// The compiler walks the class table in declared order (which is also
/// inheritance order — parents declared before children), generates one
/// constructor chunk per class, and binds:
///
/// - **setters** for every property in `properties` — compiled bytecode
///   chunks (`build_setter_chunk`) that call `vybe:gui::controlSetProperty`
/// - **methods** for every entry in `methods` — compiled thunk chunks
///   (`build_method_thunk_chunk`) that forward `(this, ...args)` to a host
///   import
///
/// Concrete leaves (those with `widget_host_fn = Some(_)`) also call the
/// host fn at the end of their ctor and copy the resulting backing
/// object's identity fields (`__control_name`, `__control_type`) onto
/// `this`. The inherited setters and methods stay intact.
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

    /// Methods added at this class level (NOT inherited). Each method is
    /// bound on the instance under its lowercased name as a thunk that
    /// forwards `(this, ...args)` to the corresponding host fn import.
    ///
    /// Real .NET semantics: methods are virtual unless marked `sealed`,
    /// and the most-derived binding wins. Our chain registers parents
    /// before children, so a child re-declaring a method with the same
    /// name overwrites the parent's binding via `struct_set` — same
    /// dispatch shape as virtual override.
    pub methods: &'static [DotnetMethod],

    /// User-visible constructor arity (excluding the implicit `this`).
    /// Most classes are arity-0 (`Form`, `Button`, `Label`, …) — the user
    /// subclasses them and writes their own ctor that calls
    /// `MyBase.New()`. Value-shaped types (`Pen`, `SolidBrush`, `Bitmap`,
    /// `Font`) take args directly: `New Pen(Color.Red, 5)`.
    ///
    /// When `ctor_arity > 0`, the args are passed straight through to
    /// `widget_host_fn` — they're NOT forwarded to the parent ctor, which
    /// is always called with 0 args (abstract bases like `Object`,
    /// `Component`, `Brush` take no args at this layer).
    pub ctor_arity: u8,

    /// If `Some(host_fn)`, this is a concrete leaf class — its ctor calls
    /// `<widget_host_module>::<host_fn>` at the end of construction to
    /// wire the underlying backing object. For arity-0 classes the host
    /// fn is called with 0 args; for arity-N classes the user-supplied
    /// ctor args are forwarded.
    ///
    /// Abstract bases (`Object`, `Control`, `ButtonBase`, `TextBoxBase`,
    /// `ListControl`, `ScrollableControl`, `ContainerControl`, `Brush`, …)
    /// leave this `None` — they bind setters/methods but don't materialize
    /// a backing object themselves.
    pub widget_host_fn: Option<&'static str>,

    /// Host module for `widget_host_fn`. `"vybe:gui"` for all WinForms
    /// controls and drawing types (`Pen`/`SolidBrush`/`Graphics`/etc.).
    /// Defaulted to `"vybe:gui"` in family files via the
    /// [`gui_class!`] / explicit-module convention so the GUI control
    /// families don't have to repeat the module everywhere.
    pub widget_host_module: &'static str,
}

impl DotnetClass {
    /// `true` for value-type-style classes whose ctor should return
    /// the backing host object **directly** rather than building a
    /// fresh `this` + copying identity fields. Used by
    /// `System.Drawing.Point` / `System.Drawing.Size` — the host fn
    /// already produces the exact `{x, y}` / `{width, height}` shape
    /// the GUI dispatch reads, so introducing an intermediate `this`
    /// object strips those fields and leaves controls stacked at (0,0).
    pub fn is_value_type(&self) -> bool {
        self.widget_host_fn.is_some() && self.parent.is_none() && self.methods.is_empty()
    }
}

/// One method on a `DotnetClass`.
///
/// At registration time the compiler builds a thunk chunk per method
/// (see `builder::build_method_thunk_chunk`) that bridges between the
/// VM-level method call and the `target`'s underlying implementation.
/// The thunk is bound on the class instance under `name.to_lowercase()`.
/// When user code does `obj.MethodName(args)` the VM does
/// `struct_get "methodname"` → method ref → call with `this` first.
#[derive(Debug, Clone, Copy)]
pub struct DotnetMethod {
    /// PascalCase method name (`"DrawLine"`, `"CreateGraphics"`). Bound
    /// on the instance under `name.to_lowercase()` to match the canonical
    /// AST.
    pub name: &'static str,

    /// Total arity including the implicit `this`. A no-arg method like
    /// `Show()` has `arity = 1`; `DrawLine(p, x1, y1, x2, y2)` has
    /// `arity = 6`.
    pub arity: u8,

    /// Where the method's implementation lives.
    pub target: MethodTarget,
}

/// What a `DotnetMethod` thunk forwards to.
///
/// Three cases:
///
/// - **`Host`** — the method's implementation is a registered host fn.
///   The thunk does `call_import <module>::<fn>` with `(this, arg0, ...)`
///   on the stack. The common case for one-shot forwarders (`Show`,
///   `Hide`, `__ctrl_show`, …).
///
/// - **`DotnetCtor`** — the method returns a fresh instance of another
///   .NET class. The thunk does `global_get <class> ; call(N-1)` —
///   passing the user args (NOT `this`) to the target class's ctor.
///   Used for factory-style methods that DON'T need `this`.
///
/// - **`Body`** — a small declarative bytecode template ([`MethodOp`]
///   sequence) the builder compiles to real opcodes. Used wherever the
///   .NET API doesn't map 1:1 onto a single host fn — e.g.
///   `Graphics.DrawLine(pen, x1, y1, x2, y2)` reads `pen.color` /
///   `pen.width` from `this`'s args, then makes ~5 sequential canvas
///   host calls. The body is a static slice that lives next to the
///   class definition; the builder handles all the lowering. There's
///   no interpreter at runtime — `Body` compiles to identical bytecode
///   to what hand-written ops would emit.
#[derive(Debug, Clone, Copy)]
pub enum MethodTarget {
    Host {
        module: &'static str,
        fn_name: &'static str,
    },
    DotnetCtor {
        class: &'static str,
    },
    Body(&'static [MethodOp]),
}

impl MethodTarget {
    /// Convenience constructor for `Host` variant.
    pub const fn host(module: &'static str, fn_name: &'static str) -> Self {
        MethodTarget::Host { module, fn_name }
    }

    /// Convenience constructor for `DotnetCtor` variant.
    pub const fn dotnet_ctor(class: &'static str) -> Self {
        MethodTarget::DotnetCtor { class }
    }

    /// Convenience constructor for `Body` variant.
    pub const fn body(ops: &'static [MethodOp]) -> Self {
        MethodTarget::Body(ops)
    }

    /// True if this body is a pure no-op (`return null`) — e.g.
    /// `SuspendLayout`/`ResumeLayout`/`PerformLayout`. These resolve through
    /// the profile's `noop` value-method, so no emitted thunk is needed.
    pub fn is_noop(&self) -> bool {
        matches!(
            self,
            MethodTarget::Body([MethodOp::PushConstNull, MethodOp::Return])
        )
    }
}

/// One operation in a [`MethodTarget::Body`] template.
///
/// `Body` sequences are small declarative bytecode templates: each op
/// lowers to one or two real opcodes via
/// [`super::builder::build_method_thunk_chunk`]. The lowering is
/// mechanical — the same opcodes you'd write by hand, just generated
/// from the slice.
///
/// ## Stack discipline
///
/// Each op documents its stack effect. The builder doesn't statically
/// verify them — sequences are written by hand and tested via
/// integration. Standard convention:
///
/// - `Push*` ops add to the stack
/// - `CallHost` / `NewDotnet` consume their args and leave the result
/// - `Drop` removes the top of the stack
/// - `SetField` consumes `[obj, val]` and leaves nothing
/// - `Return` returns top-of-stack (or null if the stack is empty)
///
/// Method args are 1-indexed: arg `1` is the first user-supplied arg
/// AFTER `this`. So `Graphics.DrawLine(pen, x1, y1, x2, y2)` has `pen`
/// at `PushArg(1)`, `x1` at `PushArg(2)`, etc.
#[derive(Debug, Clone, Copy)]
pub enum MethodOp {
    /// Push `this` (slot 0 in the call frame — WASM convention: the first
    /// argument is slot 0, and `this` is passed as the first argument).
    PushThis,
    /// Push user arg `n` (1-indexed). Arg 1 is the first user arg
    /// AFTER `this`.
    PushArg(u8),
    /// Push `this.<field>`.
    PushThisField(&'static str),
    /// Push `argN.<field>`.
    PushArgField(u8, &'static str),
    /// Push `argN.<f1>.<f2>` (two-level field access). Convenience op
    /// for the common case of reading a sub-field on a value-type
    /// argument — e.g. `pen.color.r` from `Graphics.DrawLine(pen, ...)`.
    /// Equivalent to `PushArgField(n, f1)` followed by a struct_get on
    /// `f2`, but expressed as a single declarative op.
    PushArgFieldField(u8, &'static str, &'static str),
    /// Push a constant integer.
    PushConstInt(i32),
    /// Push a constant float.
    PushConstFloat(f64),
    /// Push a constant string.
    PushConstStr(&'static str),
    /// Push a constant boolean.
    PushConstBool(bool),
    /// Push the `null` constant.
    PushConstNull,
    /// Call `<module>::<fn_name>` with `argc` arguments popped from
    /// the stack. Result is left on the stack.
    CallHost {
        module: &'static str,
        fn_name: &'static str,
        argc: u8,
    },
    /// Call the dotnet class `class`'s ctor with `argc` arguments
    /// popped from the stack (no implicit `this`). The class's
    /// constructor global must already be installed by an earlier
    /// `register_dotnet_classes` iteration. Result is left on the stack.
    NewDotnet { class: &'static str, argc: u8 },
    /// `struct_set` — pops `[obj, val]`, stores `val` into `obj.<field>`,
    /// leaves the stored value on the stack. (Mirrors the existing
    /// VM `struct_set` semantics.)
    SetField(&'static str),
    /// Drop top of stack.
    Drop,
    /// Duplicate top of stack.
    Dup,
    /// Return top of stack. If the stack is empty, returns null.
    Return,
}

impl DotnetClass {
    /// True if this class has a backing host constructor.
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
        v.extend_from_slice(drawing::classes());
        v
    });
    &TABLE
}
