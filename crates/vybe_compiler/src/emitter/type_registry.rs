//! Compile-time .NET type hierarchy.
//!
//! Single source of truth for type identity and inheritance across all compilers.
//! Replaces the scattered string matching (`starts_with("system.")`,
//! `matches!(name, "form" | "control" | ...)`) with proper chain walks.
//!
//! ## Architecture
//!
//! - **Framework types** — ~100 .NET BCL types with parent chains, stored in a
//!   static `LazyLock<FrameworkTypeTable>`. Read-only after init.
//! - **User-defined types** — registered per compilation in `CompileTimeTypes`.
//!   Parent references can cross into framework types (e.g. `MyForm : Form`).
//!
//! ## Resolution
//!
//! Names are always lowercased internally. Framework types can be referenced by
//! short name (`"form"`) or FQN (`"system.windows.forms.form"`). User types are
//! stored under their lowercased name.
//!
//! `is_subtype_of(child, ancestor)` walks the chain through both user and
//! framework types, with a depth limit of 50 to prevent infinite loops.

use std::collections::{HashMap, HashSet};
use std::sync::LazyLock;

// ── Public API: per-compilation type context ─────────────────────────────────

/// Per-compilation type context. Holds user-defined types and delegates
/// framework type lookups to the static table.
///
/// Add one to each compiler's struct and call `register_user_type` as classes
/// are compiled. All query methods (`is_subtype_of`, `is_control_type`, etc.)
/// walk both user and framework chains.
pub struct CompileTimeTypes {
    /// User-defined types: lowercase name → parent name (lowercased, or None)
    user_parents: HashMap<String, Option<String>>,
    /// User-defined type fields (own + inherited, accumulated at registration)
    user_fields: HashMap<String, HashSet<String>>,
    /// User-defined type methods (own + inherited, accumulated at registration)
    user_methods: HashMap<String, HashSet<String>>,
}

impl CompileTimeTypes {
    pub fn new() -> Self {
        CompileTimeTypes {
            user_parents: HashMap::new(),
            user_fields: HashMap::new(),
            user_methods: HashMap::new(),
        }
    }

    // ── Registration ─────────────────────────────────────────────────────

    /// Register a user-defined class with its parent and members.
    /// `parent` is the raw type name from the AST (short or FQN).
    /// `fields` and `methods` should include inherited members (already
    /// accumulated by the compiler's class_field_map/class_method_map).
    pub fn register_user_type(
        &mut self,
        name: &str,
        parent: Option<&str>,
        fields: HashSet<String>,
        methods: HashSet<String>,
    ) {
        let name_lower = name.to_lowercase();
        let parent_lower = parent.map(|p| p.to_lowercase());
        self.user_parents.insert(name_lower.clone(), parent_lower);
        self.user_fields.insert(name_lower.clone(), fields);
        self.user_methods.insert(name_lower, methods);
    }

    // ── Queries ──────────────────────────────────────────────────────────

    /// Is this a known .NET framework type?
    /// Accepts both short names ("Form") and FQNs ("System.Windows.Forms.Form").
    pub fn is_framework_type(&self, name: &str) -> bool {
        FRAMEWORK.resolve(&name.to_lowercase()).is_some()
    }

    /// Is `child` a subtype of (or equal to) `ancestor`?
    /// Walks the full chain through user types → framework types.
    /// Accepts short names or FQNs for both arguments.
    pub fn is_subtype_of(&self, child: &str, ancestor: &str) -> bool {
        let child_norm = self.normalize(child);
        let ancestor_norm = self.normalize(ancestor);
        if child_norm == ancestor_norm {
            return true;
        }

        let mut current = Some(child_norm);
        for _ in 0..50 {
            let name = match current {
                Some(ref n) => n.clone(),
                None => return false,
            };

            // Try user type parent first
            if let Some(parent_opt) = self.user_parents.get(&name) {
                match parent_opt {
                    Some(p) => {
                        let p_norm = self.normalize(p);
                        if p_norm == ancestor_norm {
                            return true;
                        }
                        current = Some(p_norm);
                    }
                    None => return false,
                }
            }
            // Then framework type parent
            else if let Some(parent_fqn) = FRAMEWORK.parent(&name) {
                if parent_fqn == ancestor_norm {
                    return true;
                }
                current = Some(parent_fqn);
            } else {
                return false;
            }
        }
        false
    }

    /// Does this type derive from System.Windows.Forms.Control or
    /// System.ComponentModel.Component?
    /// These are the types whose derived classes use the InitializeComponent pattern.
    pub fn is_control_type(&self, name: &str) -> bool {
        self.is_subtype_of(name, "control")
            || self.is_subtype_of(name, "component")
    }

    /// Does this type derive from System.Windows.Forms.Form?
    pub fn is_form_type(&self, name: &str) -> bool {
        self.is_subtype_of(name, "form")
    }

    /// Does this type have a user-defined parent (not a framework base)?
    /// Used to decide: call parent constructor (user) vs create new object (framework).
    pub fn has_user_parent(&self, name: &str) -> bool {
        let lower = name.to_lowercase();
        if let Some(Some(parent)) = self.user_parents.get(&lower) {
            // Parent exists — is it a user type?
            self.user_parents.contains_key(parent)
        } else {
            false
        }
    }

    /// Get accumulated fields for a user type (own + inherited).
    pub fn user_fields(&self, name: &str) -> Option<&HashSet<String>> {
        self.user_fields.get(&name.to_lowercase())
    }

    /// Get accumulated methods for a user type (own + inherited).
    pub fn user_methods(&self, name: &str) -> Option<&HashSet<String>> {
        self.user_methods.get(&name.to_lowercase())
    }

    /// Is this name a known user-defined type?
    pub fn is_user_type(&self, name: &str) -> bool {
        self.user_parents.contains_key(&name.to_lowercase())
    }

    // ── Internal ─────────────────────────────────────────────────────────

    /// Normalize a type name: resolve short → FQN for framework types.
    fn normalize(&self, name: &str) -> String {
        let lower = name.to_lowercase();
        FRAMEWORK.resolve(&lower).unwrap_or(lower)
    }
}

// ── Static framework type table ──────────────────────────────────────────────

/// Read-only .NET framework type hierarchy. Loaded once via LazyLock.
/// Contains parent chains for ~100 commonly used BCL types.
struct FrameworkTypeTable {
    /// FQN (lowercased) → parent FQN (lowercased). None for System.Object.
    parents: HashMap<String, Option<String>>,
    /// Short name (lowercased) → FQN (lowercased).
    /// First registration wins (e.g. "timer" → "system.windows.forms.timer").
    short_to_fqn: HashMap<String, String>,
}

impl FrameworkTypeTable {
    fn new() -> Self {
        let mut table = FrameworkTypeTable {
            parents: HashMap::new(),
            short_to_fqn: HashMap::new(),
        };
        table.load();
        table
    }

    fn add(&mut self, fqn: &str, parent: Option<&str>) {
        let fqn_lower = fqn.to_lowercase();
        let short = fqn_lower.rsplit('.').next().unwrap_or(&fqn_lower).to_string();
        let parent_lower = parent.map(|p| p.to_lowercase());
        self.parents.insert(fqn_lower.clone(), parent_lower);
        self.short_to_fqn.entry(short).or_insert(fqn_lower);
    }

    /// Resolve a name (short or FQN, already lowercased) to its FQN.
    /// Returns None if not a framework type.
    fn resolve(&self, lower_name: &str) -> Option<String> {
        // Direct FQN match
        if self.parents.contains_key(lower_name) {
            return Some(lower_name.to_string());
        }
        // Short name lookup
        self.short_to_fqn.get(lower_name).cloned()
    }

    /// Get the parent FQN for a type (by FQN or short name, already lowercased).
    fn parent(&self, lower_name: &str) -> Option<String> {
        // Direct FQN lookup
        if let Some(parent) = self.parents.get(lower_name) {
            return parent.clone();
        }
        // Short name → FQN → parent
        if let Some(fqn) = self.short_to_fqn.get(lower_name) {
            if let Some(parent) = self.parents.get(fqn) {
                return parent.clone();
            }
        }
        None
    }

    fn load(&mut self) {
        // ── Core types ───────────────────────────────────────────────────
        self.add("System.Object", None);
        self.add("System.ValueType", Some("System.Object"));
        self.add("System.Enum", Some("System.ValueType"));
        self.add("System.String", Some("System.Object"));
        self.add("System.Math", Some("System.Object"));
        self.add("System.Convert", Some("System.Object"));
        self.add("System.DateTime", Some("System.ValueType"));
        self.add("System.TimeSpan", Some("System.ValueType"));
        self.add("System.Guid", Some("System.ValueType"));
        self.add("System.EventArgs", Some("System.Object"));
        self.add("System.Exception", Some("System.Object"));
        self.add("System.Random", Some("System.Object"));

        // ── IO ───────────────────────────────────────────────────────────
        self.add("System.IO.Stream", Some("System.Object"));
        self.add("System.IO.StreamReader", Some("System.Object"));
        self.add("System.IO.StreamWriter", Some("System.Object"));
        self.add("System.IO.FileStream", Some("System.IO.Stream"));
        self.add("System.IO.MemoryStream", Some("System.IO.Stream"));
        self.add("System.IO.BinaryReader", Some("System.Object"));
        self.add("System.IO.BinaryWriter", Some("System.Object"));

        // ── Collections ──────────────────────────────────────────────────
        self.add("System.Collections.ArrayList", Some("System.Object"));
        self.add("System.Collections.Hashtable", Some("System.Object"));
        self.add("System.Collections.SortedList", Some("System.Object"));
        self.add("System.Collections.Generic.List", Some("System.Object"));
        self.add("System.Collections.Generic.Dictionary", Some("System.Object"));
        self.add("System.Collections.Generic.Queue", Some("System.Object"));
        self.add("System.Collections.Generic.Stack", Some("System.Object"));
        self.add("System.Collections.Generic.HashSet", Some("System.Object"));

        // ── Text ─────────────────────────────────────────────────────────
        self.add("System.Text.StringBuilder", Some("System.Object"));
        self.add("System.Text.RegularExpressions.Regex", Some("System.Object"));

        // ── Threading ────────────────────────────────────────────────────
        self.add("System.Threading.Thread", Some("System.Object"));
        self.add("System.Threading.Tasks.Task", Some("System.Object"));
        self.add("System.Threading.Timer", Some("System.Object"));
        self.add("System.Threading.Mutex", Some("System.Object"));
        self.add("System.Threading.Semaphore", Some("System.Object"));

        // ── Diagnostics ──────────────────────────────────────────────────
        self.add("System.Diagnostics.Stopwatch", Some("System.Object"));
        self.add("System.Diagnostics.Process", Some("System.Object"));
        self.add("System.Diagnostics.Debug", Some("System.Object"));
        self.add("System.Diagnostics.Trace", Some("System.Object"));

        // ── Drawing ──────────────────────────────────────────────────────
        self.add("System.Drawing.Point", Some("System.ValueType"));
        self.add("System.Drawing.PointF", Some("System.ValueType"));
        self.add("System.Drawing.Size", Some("System.ValueType"));
        self.add("System.Drawing.SizeF", Some("System.ValueType"));
        self.add("System.Drawing.Color", Some("System.ValueType"));
        self.add("System.Drawing.Font", Some("System.Object"));
        self.add("System.Drawing.Pen", Some("System.Object"));
        self.add("System.Drawing.SolidBrush", Some("System.Object"));
        self.add("System.Drawing.Graphics", Some("System.Object"));
        self.add("System.Drawing.Bitmap", Some("System.Object"));
        self.add("System.Drawing.Image", Some("System.Object"));

        // ── WinForms hierarchy (correct .NET chain) ──────────────────────
        // Form → ContainerControl → ScrollableControl → Control → Component → MarshalByRefObject → Object
        self.add("System.MarshalByRefObject", Some("System.Object"));
        self.add("System.ComponentModel.Component", Some("System.MarshalByRefObject"));
        self.add("System.Windows.Forms.Control", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.ScrollableControl", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.ContainerControl", Some("System.Windows.Forms.ScrollableControl"));
        self.add("System.Windows.Forms.Form", Some("System.Windows.Forms.ContainerControl"));
        self.add("System.Windows.Forms.UserControl", Some("System.Windows.Forms.ContainerControl"));

        // ── Controls ─────────────────────────────────────────────────────
        self.add("System.Windows.Forms.ButtonBase", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.Button", Some("System.Windows.Forms.ButtonBase"));
        self.add("System.Windows.Forms.Label", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.TextBoxBase", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.TextBox", Some("System.Windows.Forms.TextBoxBase"));
        self.add("System.Windows.Forms.RichTextBox", Some("System.Windows.Forms.TextBoxBase"));
        self.add("System.Windows.Forms.MaskedTextBox", Some("System.Windows.Forms.TextBoxBase"));
        self.add("System.Windows.Forms.CheckBox", Some("System.Windows.Forms.ButtonBase"));
        self.add("System.Windows.Forms.RadioButton", Some("System.Windows.Forms.ButtonBase"));
        self.add("System.Windows.Forms.ListControl", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.ComboBox", Some("System.Windows.Forms.ListControl"));
        self.add("System.Windows.Forms.ListBox", Some("System.Windows.Forms.ListControl"));
        self.add("System.Windows.Forms.Panel", Some("System.Windows.Forms.ScrollableControl"));
        self.add("System.Windows.Forms.GroupBox", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.TabControl", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.TabPage", Some("System.Windows.Forms.Panel"));
        self.add("System.Windows.Forms.DataGridView", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.ProgressBar", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.TrackBar", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.NumericUpDown", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.DateTimePicker", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.PictureBox", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.ToolStrip", Some("System.Windows.Forms.ScrollableControl"));
        self.add("System.Windows.Forms.MenuStrip", Some("System.Windows.Forms.ToolStrip"));
        self.add("System.Windows.Forms.StatusStrip", Some("System.Windows.Forms.ToolStrip"));
        self.add("System.Windows.Forms.SplitContainer", Some("System.Windows.Forms.ContainerControl"));
        self.add("System.Windows.Forms.FlowLayoutPanel", Some("System.Windows.Forms.Panel"));
        self.add("System.Windows.Forms.TableLayoutPanel", Some("System.Windows.Forms.Panel"));
        self.add("System.Windows.Forms.LinkLabel", Some("System.Windows.Forms.Label"));
        self.add("System.Windows.Forms.TreeView", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.ListView", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.WebBrowser", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.MonthCalendar", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.HScrollBar", Some("System.Windows.Forms.Control"));
        self.add("System.Windows.Forms.VScrollBar", Some("System.Windows.Forms.Control"));

        // ── Non-visual components ────────────────────────────────────────
        self.add("System.Windows.Forms.Timer", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.ToolTip", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.ImageList", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.BindingSource", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.ErrorProvider", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.HelpProvider", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.BackgroundWorker", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.NotifyIcon", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.BindingNavigator", Some("System.Windows.Forms.ToolStrip"));

        // ── Dialogs ──────────────────────────────────────────────────────
        self.add("System.Windows.Forms.CommonDialog", Some("System.ComponentModel.Component"));
        self.add("System.Windows.Forms.FileDialog", Some("System.Windows.Forms.CommonDialog"));
        self.add("System.Windows.Forms.OpenFileDialog", Some("System.Windows.Forms.FileDialog"));
        self.add("System.Windows.Forms.SaveFileDialog", Some("System.Windows.Forms.FileDialog"));
        self.add("System.Windows.Forms.FolderBrowserDialog", Some("System.Windows.Forms.CommonDialog"));
        self.add("System.Windows.Forms.ColorDialog", Some("System.Windows.Forms.CommonDialog"));
        self.add("System.Windows.Forms.FontDialog", Some("System.Windows.Forms.CommonDialog"));

        // ── Data ─────────────────────────────────────────────────────────
        self.add("System.Data.DataTable", Some("System.Object"));
        self.add("System.Data.DataSet", Some("System.Object"));
        self.add("System.Data.DataRow", Some("System.Object"));
        self.add("System.Data.DataColumn", Some("System.Object"));
        self.add("System.Data.SqlClient.SqlConnection", Some("System.Object"));
        self.add("System.Data.SqlClient.SqlCommand", Some("System.Object"));
        self.add("System.Data.SqlClient.SqlDataReader", Some("System.Object"));
        self.add("System.Data.SqlClient.SqlDataAdapter", Some("System.Object"));
        self.add("System.Data.SqlClient.SqlTransaction", Some("System.Object"));
        self.add("System.Data.OleDb.OleDbConnection", Some("System.Object"));
        self.add("System.Data.OleDb.OleDbCommand", Some("System.Object"));
        self.add("ADODB.Connection", Some("System.Object"));
        self.add("ADODB.Command", Some("System.Object"));
        self.add("ADODB.Recordset", Some("System.Object"));

        // ── Net ──────────────────────────────────────────────────────────
        self.add("System.Net.Sockets.TcpClient", Some("System.Object"));
        self.add("System.Net.Sockets.TcpListener", Some("System.Object"));
        self.add("System.Net.Sockets.UdpClient", Some("System.Object"));
        self.add("System.Net.Sockets.Socket", Some("System.Object"));

        // ── Security ─────────────────────────────────────────────────────
        self.add("System.Security.Cryptography.HashAlgorithm", Some("System.Object"));

        // ── XML ──────────────────────────────────────────────────────────
        self.add("System.Xml.Linq.XDocument", Some("System.Object"));
        self.add("System.Xml.Linq.XElement", Some("System.Object"));
    }
}

static FRAMEWORK: LazyLock<FrameworkTypeTable> = LazyLock::new(|| FrameworkTypeTable::new());

// ── Convenience: update needs_auto_init_component ────────────────────────────

/// Check if a class should get an auto-generated `InitializeComponent()` call.
/// Uses the type registry for proper inheritance checking instead of string matching.
///
/// Returns true when:
/// 1. No explicit constructor, AND
/// 2. The class has an `initializecomponent` method, AND
/// 3. The class inherits from a Control or Component type (checked via registry)
pub fn needs_auto_init_component(
    has_explicit_ctor: bool,
    method_names: &[String],
    base_type: Option<&str>,
    types: &CompileTimeTypes,
) -> bool {
    !has_explicit_ctor
        && method_names.iter().any(|m| m.eq_ignore_ascii_case("initializecomponent"))
        && base_type.map(|b| types.is_control_type(b)).unwrap_or(false)
}
