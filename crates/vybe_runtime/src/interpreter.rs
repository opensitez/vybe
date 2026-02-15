use crate::builtins::*;
use crate::environment::Environment;
use crate::evaluator::{evaluate, values_equal, value_in_range, compare_values};
use crate::event_system::EventSystem;
use crate::value::{ExitType, RuntimeError, Value, ObjectData};
use crate::EventData;
use std::collections::{HashMap, VecDeque};
use std::io::BufRead;
use std::rc::Rc;
use std::cell::RefCell;
use std::sync::mpsc;
use vybe_parser::{CaseCondition, Declaration, Expression, FunctionDecl, Identifier, Program, Statement, SubDecl};

pub struct Interpreter {
    pub env: Environment,
    pub functions: HashMap<String, FunctionDecl>,
    pub subs: HashMap<String, SubDecl>,
    pub classes: HashMap<String, vybe_parser::ClassDecl>,
    pub events: EventSystem,
    pub side_effects: VecDeque<crate::RuntimeSideEffect>,
    current_module: Option<String>, // Track which form/module is currently executing
    current_object: Option<Rc<RefCell<crate::value::ObjectData>>>,
    with_object: Option<Value>,
    pub file_handles: HashMap<i32, crate::file_io::FileHandle>,
    pub net_handles: HashMap<i64, crate::builtins::networking::NetHandle>,
    next_net_handle_id: i64,
    pub resources: HashMap<String, String>,
    pub resource_entries: Vec<crate::ResourceEntry>,
    pub command_line_args: Vec<String>,
    /// Optional channel for sending console output to the UI (interactive console mode).
    pub console_tx: Option<mpsc::Sender<crate::ConsoleMessage>>,
    /// Optional channel for receiving console input from the UI (Console.ReadLine).
    pub console_input_rx: Option<mpsc::Receiver<String>>,
    /// When true, console output goes directly to stdout (CLI mode).
    /// When false (default), output is buffered in side_effects (tests/form mode).
    pub direct_console: bool,
    /// Active Imports statements — maps short names to fully-qualified paths.
    /// E.g. Imports System.IO → allows using "File" to resolve "System.IO.File".
    /// Also stores aliases: Imports IO = System.IO → ("IO", "System.IO").
    pub imports: Vec<ImportEntry>,
    /// Maps fully-qualified class name → key in self.classes.
    /// E.g. "myapp.models.customer" → "customer" (the actual key in self.classes).
    pub namespace_map: HashMap<String, String>,
    /// Registered interfaces (name → InterfaceDecl).
    pub interfaces: HashMap<String, vybe_parser::InterfaceDecl>,
    /// Registered structures (treated like value-type classes).
    pub structures: HashMap<String, vybe_parser::StructureDecl>,
    /// Registered delegates (name → DelegateDecl).
    pub delegates: HashMap<String, vybe_parser::DelegateDecl>,
    /// On Error Resume Next active?
    pub on_error_resume_next: bool,
    /// On Error GoTo <label> — the label to jump to on error (None = disabled).
    pub on_error_goto_label: Option<String>,
    /// Static local variables: key = "module.proc.var_name" → Value.
    pub static_locals: HashMap<String, Value>,
    /// Track which Sub/Function is currently executing (for static locals).
    current_procedure: Option<String>,
}

/// An active Imports entry.
#[derive(Debug, Clone)]
pub struct ImportEntry {
    /// The full path, e.g. "System.IO"
    pub path: String,
    /// Optional alias, e.g. Some("IO") for `Imports IO = System.IO`
    pub alias: Option<String>,
}

struct RuntimeRegistry {
    threads: HashMap<String, std::thread::JoinHandle<()>>,
    processes: HashMap<String, std::process::Child>,
    shared_objects: HashMap<String, std::sync::Arc<std::sync::Mutex<crate::value::SharedObjectData>>>,
}

static REGISTRY: std::sync::OnceLock<std::sync::Mutex<RuntimeRegistry>> = std::sync::OnceLock::new();

fn get_registry() -> &'static std::sync::Mutex<RuntimeRegistry> {
    REGISTRY.get_or_init(|| std::sync::Mutex::new(RuntimeRegistry {
        threads: HashMap::new(),
        processes: HashMap::new(),
        shared_objects: HashMap::new(),
    }))
}

fn generate_runtime_id() -> String {
    use std::sync::atomic::{AtomicUsize, Ordering};
    static COUNTER: AtomicUsize = AtomicUsize::new(1);
    format!("rt_{}", COUNTER.fetch_add(1, Ordering::SeqCst))
}

impl Interpreter {
    pub fn new() -> Self {
        let mut interp = Self {
            env: Environment::new(),
            functions: HashMap::new(),
            subs: HashMap::new(),
            classes: HashMap::new(),
            events: EventSystem::new(),
            side_effects: VecDeque::new(),
            current_module: None,
            current_object: None,
            with_object: None,
            file_handles: HashMap::new(),
            net_handles: HashMap::new(),
            next_net_handle_id: 1,
            resources: HashMap::new(),
            resource_entries: Vec::new(),
            command_line_args: Vec::new(),
            console_tx: None,
            console_input_rx: None,
            direct_console: false,
            imports: Vec::new(),
            namespace_map: HashMap::new(),
            interfaces: HashMap::new(),
            structures: HashMap::new(),
            delegates: HashMap::new(),
            on_error_resume_next: false,
            on_error_goto_label: None,
            static_locals: HashMap::new(),
            current_procedure: None,
        };
        interp.register_builtin_constants();
        interp.init_namespaces();
        interp
    }

    pub fn new_background(
        functions: HashMap<String, vybe_parser::ast::decl::FunctionDecl>,
        subs: HashMap<String, vybe_parser::ast::decl::SubDecl>,
        classes: HashMap<String, vybe_parser::ast::decl::ClassDecl>,
        namespace_map: HashMap<String, String>,
    ) -> Self {
        let mut interp = Self::new();
        interp.functions = functions;
        interp.subs = subs;
        interp.classes = classes;
        interp.namespace_map = namespace_map;
        interp
    }

    pub fn init_namespaces(&mut self) {
        // Create System.IO.File object
        let file_obj_data = ObjectData {
            class_name: "System.IO.File".to_string(),
            fields: HashMap::new(),
        };
        let file_obj = Value::Object(Rc::new(RefCell::new(file_obj_data)));

        // Create System.IO.Path object
        let path_obj_data = ObjectData {
            class_name: "System.IO.Path".to_string(),
            fields: HashMap::new(),
        };
        let path_obj = Value::Object(Rc::new(RefCell::new(path_obj_data)));

        // Create System.Console object with default properties
        let mut console_fields = HashMap::new();
        console_fields.insert("foregroundcolor".to_string(), Value::Integer(7));  // Gray
        console_fields.insert("backgroundcolor".to_string(), Value::Integer(0)); // Black
        console_fields.insert("title".to_string(), Value::String(String::new()));
        console_fields.insert("cursorleft".to_string(), Value::Integer(0));
        console_fields.insert("cursortop".to_string(), Value::Integer(0));
        console_fields.insert("cursorvisible".to_string(), Value::Boolean(true));
        console_fields.insert("windowwidth".to_string(), Value::Integer(80));
        console_fields.insert("windowheight".to_string(), Value::Integer(25));
        console_fields.insert("bufferwidth".to_string(), Value::Integer(80));
        console_fields.insert("bufferheight".to_string(), Value::Integer(300));
        console_fields.insert("keyavailable".to_string(), Value::Boolean(false));
        let console_obj_data = ObjectData {
            class_name: "System.Console".to_string(),
            fields: console_fields,
        };
        let console_obj = Value::Object(Rc::new(RefCell::new(console_obj_data)));

        // Create ConsoleColor "enum" with named color constants
        let mut cc_fields = HashMap::new();
        cc_fields.insert("black".to_string(), Value::Integer(0));
        cc_fields.insert("darkblue".to_string(), Value::Integer(1));
        cc_fields.insert("darkgreen".to_string(), Value::Integer(2));
        cc_fields.insert("darkcyan".to_string(), Value::Integer(3));
        cc_fields.insert("darkred".to_string(), Value::Integer(4));
        cc_fields.insert("darkmagenta".to_string(), Value::Integer(5));
        cc_fields.insert("darkyellow".to_string(), Value::Integer(6));
        cc_fields.insert("gray".to_string(), Value::Integer(7));
        cc_fields.insert("darkgray".to_string(), Value::Integer(8));
        cc_fields.insert("blue".to_string(), Value::Integer(9));
        cc_fields.insert("green".to_string(), Value::Integer(10));
        cc_fields.insert("cyan".to_string(), Value::Integer(11));
        cc_fields.insert("red".to_string(), Value::Integer(12));
        cc_fields.insert("magenta".to_string(), Value::Integer(13));
        cc_fields.insert("yellow".to_string(), Value::Integer(14));
        cc_fields.insert("white".to_string(), Value::Integer(15));
        let cc_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "ConsoleColor".to_string(), fields: cc_fields,
        })));

        // Create System.Math object
        let math_obj_data = ObjectData {
            class_name: "System.Math".to_string(),
            fields: HashMap::new(),
        };
        let math_obj = Value::Object(Rc::new(RefCell::new(math_obj_data)));

        // Create System.IO namespace object
        let mut io_fields = HashMap::new();
        io_fields.insert("file".to_string(), file_obj);
        io_fields.insert("path".to_string(), path_obj);
        
        let io_obj_data = ObjectData {
            class_name: "Namespace".to_string(),
            fields: io_fields,
        };
        let io_obj = Value::Object(Rc::new(RefCell::new(io_obj_data)));

        // Create System namespace object
        let mut system_fields = HashMap::new();
        system_fields.insert("io".to_string(), io_obj.clone());
        system_fields.insert("console".to_string(), console_obj.clone());
        system_fields.insert("math".to_string(), math_obj.clone());
        // System.DBNull
        let mut dbnull_fields = HashMap::new();
        dbnull_fields.insert("value".to_string(), Value::Nothing);
        let dbnull_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "DBNull".to_string(), fields: dbnull_fields,
        })));
        system_fields.insert("dbnull".to_string(), dbnull_obj.clone());
        
        // System.BitConverter
        let bit_converter_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "BitConverter".to_string(),
            fields: HashMap::new(),
        })));
        system_fields.insert("bitconverter".to_string(), bit_converter_obj);
        
        // Create System.Text.Encoding.UTF8 object
        let utf8_obj_data = ObjectData {
            class_name: "Utf8Encoding".to_string(),
            fields: HashMap::new(),
        };
        let utf8_obj = Value::Object(Rc::new(RefCell::new(utf8_obj_data)));

        let mut text_fields = HashMap::new();
        let mut encoding_fields = HashMap::new();
        encoding_fields.insert("utf8".to_string(), utf8_obj);
        let encoding_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Namespace".to_string(), fields: encoding_fields,
        })));
        text_fields.insert("encoding".to_string(), encoding_obj.clone());
        let text_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Namespace".to_string(), fields: text_fields,
        })));
        // Create System.Security.Cryptography objects
        let md5_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "MD5".to_string(), fields: HashMap::new(),
        })));
        let sha256_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "SHA256".to_string(), fields: HashMap::new(),
        })));
        
        let mut crypto_fields = HashMap::new();
        crypto_fields.insert("md5".to_string(), md5_class_obj.clone());
        crypto_fields.insert("sha256".to_string(), sha256_class_obj.clone());
        let crypto_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Namespace".to_string(), fields: crypto_fields,
        })));
        
        let mut security_fields = HashMap::new();
        security_fields.insert("cryptography".to_string(), crypto_obj.clone());
        let security_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Namespace".to_string(), fields: security_fields,
        })));
        system_fields.insert("security".to_string(), security_obj.clone());
        system_fields.insert("text".to_string(), text_obj.clone());

        // System.Diagnostics.Process + ProcessStartInfo
        let mut diag_fields = HashMap::new();
        let process_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Process".to_string(), fields: HashMap::new(),
        })));
        let psi_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "ProcessStartInfo".to_string(), fields: HashMap::new(),
        })));
        diag_fields.insert("process".to_string(), process_class_obj);
        diag_fields.insert("processstartinfo".to_string(), psi_class_obj);
        let diag_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Namespace".to_string(), fields: diag_fields,
        })));
        system_fields.insert("diagnostics".to_string(), diag_obj.clone());
        
        let system_obj_data = ObjectData {
            class_name: "Namespace".to_string(),
            fields: system_fields,
        };
        let system_obj = Value::Object(Rc::new(RefCell::new(system_obj_data)));

        // Create System.Windows.Forms namespace
        let mut swf_fields = HashMap::new();
        let app_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Application".to_string(), fields: HashMap::new(),
        })));
        swf_fields.insert("application".to_string(), app_class_obj.clone());
        
        let swf_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Namespace".to_string(), fields: swf_fields,
        })));
        
        // Note: we don't attach windows to system fields here because system fields were already created above.
        // But we DO need to make it accessible via System.Windows.Forms
        // Ideally we should modify system_fields BEFORE creating system_obj, but system_obj is created above.
        // Let's just register it in the environment directly.
        // To be correct, we should insert it into system_fields if possible, but system_fields is moved.
        // So we will just define it globally.

        // Register all nested namespaces in the environment for easy resolution by Imports
        self.env.define("system", system_obj.clone());
        self.env.define("system.io", io_obj.clone());
        self.env.define("system.text", text_obj.clone());
        self.env.define("system.text.encoding", encoding_obj.clone());
        self.env.define("system.security", security_obj.clone());
        self.env.define("system.security.cryptography", crypto_obj.clone());
        self.env.define("system.diagnostics", diag_obj.clone());
        self.env.define("system.windows.forms", swf_obj.clone());
        // Also expose Application globally for convenience (as per standard VB project imports)
        self.env.define("application", app_class_obj.clone());
        
        // Also register Console and Math globally for convenience (like implicit Imports System)
        self.env.define("console", console_obj);
        self.env.define("math", math_obj);
        // Also register DBNull globally
        self.env.define("dbnull", dbnull_obj);
        // Register ConsoleColor enum globally
        self.env.define("consolecolor", cc_obj);

        // ── System.Drawing ──────────────────────────────────────────────────────
        let mut drawing_fields = HashMap::new();
        // Color
        let color_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Color".to_string(), fields: HashMap::new(),
        })));
        drawing_fields.insert("color".to_string(), color_class_obj.clone());
        // Pen
        let pen_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Pen".to_string(), fields: HashMap::new(),
        })));
        drawing_fields.insert("pen".to_string(), pen_class_obj.clone());
        // SolidBrush
        let brush_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "SolidBrush".to_string(), fields: HashMap::new(),
        })));
        drawing_fields.insert("solidbrush".to_string(), brush_class_obj.clone());
        // Graphics
        let graphics_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Graphics".to_string(), fields: HashMap::new(),
        })));
        drawing_fields.insert("graphics".to_string(), graphics_class_obj.clone());

        let drawing_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "Namespace".to_string(), fields: drawing_fields,
        })));
        
        // Add Drawing to System namespace (using Compare-And-Swap style access since system_obj is Rc<RefCell>)
        if let Value::Object(sys_ref) = &system_obj {
            sys_ref.borrow_mut().fields.insert("drawing".to_string(), drawing_obj.clone());
        }
        self.env.define("system.drawing", drawing_obj.clone());
        // Register Color/Pen/Brush globally for easy access (standard imports)
        self.env.define("color", color_class_obj);
        self.env.define("pen", pen_class_obj);
        self.env.define("solidbrush", brush_class_obj);
        self.env.define("graphics", graphics_class_obj);

        // ── System.Windows.Forms.MessageBox ─────────────────────────────────────
        let msgbox_class_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "MessageBox".to_string(), fields: HashMap::new(),
        })));
        if let Value::Object(swf_ref) = &swf_obj {
            swf_ref.borrow_mut().fields.insert("messagebox".to_string(), msgbox_class_obj.clone());
        }
        // Register MessageBox globally
        self.env.define("messagebox", msgbox_class_obj.clone());
        self.env.define("system.windows.forms.messagebox", msgbox_class_obj); // Not strictly needed if swf is valid


        // ── My.Application ─────────────────────────────────────────────────────
        // Provides My.Application.Info.Title/Version, My.Application.DoEvents(), etc.
        let mut info_fields = HashMap::new();
        info_fields.insert("title".to_string(), Value::String(String::new()));
        info_fields.insert("version".to_string(), Value::String("1.0.0.0".to_string()));
        info_fields.insert("productname".to_string(), Value::String(String::new()));
        info_fields.insert("companyname".to_string(), Value::String(String::new()));
        info_fields.insert("copyright".to_string(), Value::String(String::new()));
        info_fields.insert("description".to_string(), Value::String(String::new()));
        let info_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "My.Application.Info".to_string(),
            fields: info_fields,
        })));

        let mut cmd_args_fields = HashMap::new();
        cmd_args_fields.insert("count".to_string(), Value::Integer(0));
        let cmd_args_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "CommandLineArgs".to_string(),
            fields: cmd_args_fields,
        })));

        let mut my_app_fields = HashMap::new();
        my_app_fields.insert("info".to_string(), info_obj);
        my_app_fields.insert("commandlineargs".to_string(), cmd_args_obj);
        let my_app_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "My.Application".to_string(),
            fields: my_app_fields,
        })));

        // Create the My namespace with Application; Resources added later by register_resource_entries
        let mut my_fields = HashMap::new();
        my_fields.insert("application".to_string(), my_app_obj.clone());
        let my_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "My".to_string(),
            fields: my_fields,
        })));
        self.env.define("my", my_obj);
        self.env.define("my.application", my_app_obj);
    }

    // ── Console helpers ──────────────────────────────────────────────────

    /// Read the current foreground/background colors from the Console object.
    fn get_console_colors(&self) -> (i32, i32) {
        if let Ok(Value::Object(ref obj)) = self.env.get("console") {
            let fields = obj.borrow();
            let fg = fields.fields.get("foregroundcolor")
                .and_then(|v| v.as_integer().ok())
                .unwrap_or(7);
            let bg = fields.fields.get("backgroundcolor")
                .and_then(|v| v.as_integer().ok())
                .unwrap_or(0);
            (fg, bg)
        } else {
            (7, 0)
        }
    }

    /// Map .NET ConsoleColor (0-15) to ANSI escape code.
    fn console_color_to_ansi_fg(color: i32) -> &'static str {
        match color {
            0  => "\x1b[30m",    // Black
            1  => "\x1b[34m",    // DarkBlue
            2  => "\x1b[32m",    // DarkGreen
            3  => "\x1b[36m",    // DarkCyan
            4  => "\x1b[31m",    // DarkRed
            5  => "\x1b[35m",    // DarkMagenta
            6  => "\x1b[33m",    // DarkYellow
            7  => "\x1b[37m",    // Gray
            8  => "\x1b[90m",    // DarkGray
            9  => "\x1b[94m",    // Blue
            10 => "\x1b[92m",    // Green
            11 => "\x1b[96m",    // Cyan
            12 => "\x1b[91m",    // Red
            13 => "\x1b[95m",    // Magenta
            14 => "\x1b[93m",    // Yellow
            15 => "\x1b[97m",    // White
            _  => "\x1b[37m",
        }
    }

    fn console_color_to_ansi_bg(color: i32) -> &'static str {
        match color {
            0  => "\x1b[40m",
            1  => "\x1b[44m",
            2  => "\x1b[42m",
            3  => "\x1b[46m",
            4  => "\x1b[41m",
            5  => "\x1b[45m",
            6  => "\x1b[43m",
            7  => "\x1b[47m",
            8  => "\x1b[100m",
            9  => "\x1b[104m",
            10 => "\x1b[102m",
            11 => "\x1b[106m",
            12 => "\x1b[101m",
            13 => "\x1b[105m",
            14 => "\x1b[103m",
            15 => "\x1b[107m",
            _  => "\x1b[40m",
        }
    }

    /// Send console output through the channel (with colors), directly to stdout (CLI),
    /// or buffered in side_effects (tests/form mode).
    fn send_console_output(&mut self, text: String) {
        let (fg, bg) = self.get_console_colors();
        if let Some(tx) = &self.console_tx {
            let _ = tx.send(crate::ConsoleMessage::Output { text, fg, bg });
        } else if self.direct_console {
            // CLI mode: print directly to stdout for immediate/interactive output
            use std::io::Write;
            if fg != 7 || bg != 0 {
                let ansi_fg = Self::console_color_to_ansi_fg(fg);
                let ansi_bg = Self::console_color_to_ansi_bg(bg);
                print!("{}{}{}\x1b[0m", ansi_fg, ansi_bg, text);
            } else {
                print!("{}", text);
            }
            let _ = std::io::stdout().flush();
        } else {
            // Test/form mode: buffer for later inspection
            self.side_effects.push_back(crate::RuntimeSideEffect::ConsoleOutput(text));
        }
    }

    /// Send debug output (always uses default colors).
    fn send_debug_output(&mut self, text: String) {
        if let Some(tx) = &self.console_tx {
            let _ = tx.send(crate::ConsoleMessage::Output { text, fg: 7, bg: 0 });
        } else if self.direct_console {
            use std::io::Write;
            print!("{}", text);
            let _ = std::io::stdout().flush();
        } else {
            self.side_effects.push_back(crate::RuntimeSideEffect::ConsoleOutput(text));
        }
    }

    // ── .NET-compatible EventArgs object factories ───────────────────────

    /// Create a System.EventArgs object (empty, base class for all events).
    pub fn make_event_args() -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("EventArgs".to_string()));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "EventArgs".to_string(),
            fields,
        })))
    }

    // ── System.Drawing struct factories ──────────────────────────────────

    pub fn make_point(x: i32, y: i32) -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("Point".to_string()));
        fields.insert("x".to_string(), Value::Integer(x));
        fields.insert("y".to_string(), Value::Integer(y));
        fields.insert("isempty".to_string(), Value::Boolean(x == 0 && y == 0));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "Point".to_string(), fields,
        })))
    }

    pub fn make_size(w: i32, h: i32) -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("Size".to_string()));
        fields.insert("width".to_string(), Value::Integer(w));
        fields.insert("height".to_string(), Value::Integer(h));
        fields.insert("isempty".to_string(), Value::Boolean(w == 0 && h == 0));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "Size".to_string(), fields,
        })))
    }

    pub fn make_rectangle(x: i32, y: i32, w: i32, h: i32) -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("Rectangle".to_string()));
        fields.insert("x".to_string(), Value::Integer(x));
        fields.insert("y".to_string(), Value::Integer(y));
        fields.insert("width".to_string(), Value::Integer(w));
        fields.insert("height".to_string(), Value::Integer(h));
        fields.insert("left".to_string(), Value::Integer(x));
        fields.insert("top".to_string(), Value::Integer(y));
        fields.insert("right".to_string(), Value::Integer(x + w));
        fields.insert("bottom".to_string(), Value::Integer(y + h));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "Rectangle".to_string(), fields,
        })))
    }

    /// Initialize type-specific default fields for WinForms control objects.
    /// Extracted to a separate function to keep `evaluate_expr` stack frame small.
    #[inline(never)]
    fn init_control_type_defaults(base_name: &str, fields: &mut std::collections::HashMap<String, Value>) {
        let base_lower = base_name.to_lowercase();
        match base_lower.as_str() {
            "button" => {
                fields.insert("flatstyle".to_string(), Value::String("Standard".to_string()));
                fields.insert("dialogresult".to_string(), Value::Integer(0));
                fields.insert("autosize".to_string(), Value::Boolean(false));
                fields.insert("autosizemode".to_string(), Value::String("GrowOnly".to_string()));
                fields.insert("image".to_string(), Value::Nothing);
                fields.insert("imagealign".to_string(), Value::String("MiddleCenter".to_string()));
                fields.insert("imageindex".to_string(), Value::Integer(-1));
                fields.insert("imagelist".to_string(), Value::Nothing);
                fields.insert("textalign".to_string(), Value::String("MiddleCenter".to_string()));
                fields.insert("textimagerelation".to_string(), Value::String("Overlay".to_string()));
                fields.insert("usevisualstyleback".to_string(), Value::Boolean(true));
                fields.insert("usemnemonic".to_string(), Value::Boolean(true));
                fields.insert("autoellipsis".to_string(), Value::Boolean(false));
            }
            "label" => {
                fields.insert("autosize".to_string(), Value::Boolean(true));
                fields.insert("textalign".to_string(), Value::String("TopLeft".to_string()));
                fields.insert("flatstyle".to_string(), Value::String("Standard".to_string()));
                fields.insert("borderstyle".to_string(), Value::String("None".to_string()));
                fields.insert("image".to_string(), Value::Nothing);
                fields.insert("imagealign".to_string(), Value::String("MiddleCenter".to_string()));
                fields.insert("imageindex".to_string(), Value::Integer(-1));
                fields.insert("imagelist".to_string(), Value::Nothing);
                fields.insert("autoellipsis".to_string(), Value::Boolean(false));
                fields.insert("usemnemonic".to_string(), Value::Boolean(true));
            }
            "checkbox" => {
                fields.insert("checked".to_string(), Value::Boolean(false));
                fields.insert("checkstate".to_string(), Value::String("Unchecked".to_string()));
                fields.insert("threestate".to_string(), Value::Boolean(false));
                fields.insert("autocheck".to_string(), Value::Boolean(true));
                fields.insert("checkalign".to_string(), Value::String("MiddleLeft".to_string()));
                fields.insert("textalign".to_string(), Value::String("MiddleLeft".to_string()));
                fields.insert("flatstyle".to_string(), Value::String("Standard".to_string()));
                fields.insert("appearance".to_string(), Value::String("Normal".to_string()));
                fields.insert("autosize".to_string(), Value::Boolean(true));
                fields.insert("image".to_string(), Value::Nothing);
                fields.insert("imagealign".to_string(), Value::String("MiddleCenter".to_string()));
                fields.insert("imageindex".to_string(), Value::Integer(-1));
                fields.insert("imagelist".to_string(), Value::Nothing);
            }
            "radiobutton" => {
                fields.insert("checked".to_string(), Value::Boolean(false));
                fields.insert("autocheck".to_string(), Value::Boolean(true));
                fields.insert("checkalign".to_string(), Value::String("MiddleLeft".to_string()));
                fields.insert("textalign".to_string(), Value::String("MiddleLeft".to_string()));
                fields.insert("flatstyle".to_string(), Value::String("Standard".to_string()));
                fields.insert("appearance".to_string(), Value::String("Normal".to_string()));
                fields.insert("autosize".to_string(), Value::Boolean(true));
                fields.insert("image".to_string(), Value::Nothing);
                fields.insert("imagealign".to_string(), Value::String("MiddleCenter".to_string()));
                fields.insert("imageindex".to_string(), Value::Integer(-1));
                fields.insert("imagelist".to_string(), Value::Nothing);
            }
            "groupbox" | "frame" => {
                fields.insert("flatstyle".to_string(), Value::String("Standard".to_string()));
                fields.insert("autosize".to_string(), Value::Boolean(false));
                fields.insert("autosizemode".to_string(), Value::String("GrowOnly".to_string()));
                fields.insert("padding".to_string(), Value::Integer(3));
                fields.insert("controls".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
            }
            "panel" => {
                fields.insert("borderstyle".to_string(), Value::String("None".to_string()));
                fields.insert("autosize".to_string(), Value::Boolean(false));
                fields.insert("autosizemode".to_string(), Value::String("GrowOnly".to_string()));
                fields.insert("autoscroll".to_string(), Value::Boolean(false));
                fields.insert("autoscrollminsize".to_string(), Value::Nothing);
                fields.insert("autoscrollmargin".to_string(), Value::Nothing);
                fields.insert("padding".to_string(), Value::Integer(0));
                fields.insert("controls".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
            }
            "picturebox" => {
                fields.insert("image".to_string(), Value::Nothing);
                fields.insert("imagelocation".to_string(), Value::String(String::new()));
                fields.insert("sizemode".to_string(), Value::String("Normal".to_string()));
                fields.insert("borderstyle".to_string(), Value::String("None".to_string()));
                fields.insert("errorimage".to_string(), Value::Nothing);
                fields.insert("initialimage".to_string(), Value::Nothing);
                fields.insert("waitonload".to_string(), Value::Boolean(false));
            }
            "webbrowser" => {
                fields.insert("url".to_string(), Value::Nothing);
                fields.insert("documenttext".to_string(), Value::String(String::new()));
                fields.insert("documenttitle".to_string(), Value::String(String::new()));
                fields.insert("cangoback".to_string(), Value::Boolean(false));
                fields.insert("cangoforward".to_string(), Value::Boolean(false));
                fields.insert("isbusy".to_string(), Value::Boolean(false));
                fields.insert("readystate".to_string(), Value::String("Uninitialized".to_string()));
                fields.insert("scriptErrorsSuppressed".to_lowercase(), Value::Boolean(false));
                fields.insert("allownavigation".to_string(), Value::Boolean(true));
                fields.insert("allowWebBrowserDrop".to_lowercase(), Value::Boolean(true));
                fields.insert("scrollbarsenabled".to_string(), Value::Boolean(true));
                fields.insert("webbrowsershortcutsenabled".to_string(), Value::Boolean(true));
                fields.insert("iswebbrowsercontextmenuenabled".to_string(), Value::Boolean(true));
            }
            "textbox" => {
                fields.insert("readonly".to_string(), Value::Boolean(false));
                fields.insert("multiline".to_string(), Value::Boolean(false));
                fields.insert("passwordchar".to_string(), Value::String(String::new()));
                fields.insert("maxlength".to_string(), Value::Integer(32767));
                fields.insert("scrollbars".to_string(), Value::Integer(0));
                fields.insert("wordwrap".to_string(), Value::Boolean(true));
                fields.insert("textalign".to_string(), Value::String("Left".to_string()));
                fields.insert("acceptsreturn".to_string(), Value::Boolean(false));
                fields.insert("acceptstab".to_string(), Value::Boolean(false));
                fields.insert("charactercasing".to_string(), Value::String("Normal".to_string()));
                fields.insert("selectionstart".to_string(), Value::Integer(0));
                fields.insert("selectionlength".to_string(), Value::Integer(0));
                fields.insert("selectedtext".to_string(), Value::String(String::new()));
                fields.insert("lines".to_string(), Value::Array(vec![]));
                fields.insert("modified".to_string(), Value::Boolean(false));
                fields.insert("hideselection".to_string(), Value::Boolean(true));
                fields.insert("borderstyle".to_string(), Value::String("Fixed3D".to_string()));
                fields.insert("textlength".to_string(), Value::Integer(0));
                fields.insert("useSystemPasswordChar".to_lowercase(), Value::Boolean(false));
                fields.insert("placeholdertext".to_string(), Value::String(String::new()));
            }
            "combobox" => {
                fields.insert("items".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("selectedindex".to_string(), Value::Integer(-1));
                fields.insert("selecteditem".to_string(), Value::Nothing);
                fields.insert("dropdownstyle".to_string(), Value::Integer(0));
                fields.insert("selectedvalue".to_string(), Value::Nothing);
                fields.insert("selectedtext".to_string(), Value::String(String::new()));
                fields.insert("maxdropdownitems".to_string(), Value::Integer(8));
                fields.insert("dropdownwidth".to_string(), Value::Integer(121));
                fields.insert("dropdownheight".to_string(), Value::Integer(106));
                fields.insert("maxlength".to_string(), Value::Integer(0));
                fields.insert("sorted".to_string(), Value::Boolean(false));
                fields.insert("flatstyle".to_string(), Value::String("Standard".to_string()));
                fields.insert("datasource".to_string(), Value::Nothing);
                fields.insert("displaymember".to_string(), Value::String(String::new()));
                fields.insert("valuemember".to_string(), Value::String(String::new()));
                fields.insert("autocompletemode".to_string(), Value::Integer(0));
                fields.insert("autocompletesource".to_string(), Value::Integer(0));
            }
            "listbox" | "checkedlistbox" => {
                fields.insert("items".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("selectedindex".to_string(), Value::Integer(-1));
                fields.insert("selecteditem".to_string(), Value::Nothing);
                fields.insert("selectionmode".to_string(), Value::Integer(1));
                fields.insert("selectedindices".to_string(), Value::Array(vec![]));
                fields.insert("selecteditems".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("sorted".to_string(), Value::Boolean(false));
                fields.insert("topindex".to_string(), Value::Integer(0));
                fields.insert("columnwidth".to_string(), Value::Integer(0));
                fields.insert("multicolumn".to_string(), Value::Boolean(false));
                fields.insert("horizontalscrollbar".to_string(), Value::Boolean(false));
                fields.insert("integralheight".to_string(), Value::Boolean(true));
                fields.insert("datasource".to_string(), Value::Nothing);
                fields.insert("displaymember".to_string(), Value::String(String::new()));
                fields.insert("valuemember".to_string(), Value::String(String::new()));
                // CheckedListBox extra
                fields.insert("checkeditems".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("checkonclick".to_string(), Value::Boolean(false));
            }
            "domainupdown" => {
                fields.insert("items".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("selectedindex".to_string(), Value::Integer(-1));
                fields.insert("selecteditem".to_string(), Value::Nothing);
                fields.insert("text".to_string(), Value::String(String::new()));
                fields.insert("readonly".to_string(), Value::Boolean(false));
                fields.insert("wrap".to_string(), Value::Boolean(false));
                fields.insert("sorted".to_string(), Value::Boolean(false));
            }
            "progressbar" => {
                fields.insert("value".to_string(), Value::Integer(0));
                fields.insert("minimum".to_string(), Value::Integer(0));
                fields.insert("maximum".to_string(), Value::Integer(100));
                fields.insert("step".to_string(), Value::Integer(10));
                fields.insert("style".to_string(), Value::Integer(0));
                fields.insert("marqueeanimationspeed".to_string(), Value::Integer(100));
                fields.insert("righttoleftlayout".to_string(), Value::Boolean(false));
            }
            "numericupdown" => {
                fields.insert("value".to_string(), Value::Integer(0));
                fields.insert("minimum".to_string(), Value::Integer(0));
                fields.insert("maximum".to_string(), Value::Integer(100));
                fields.insert("increment".to_string(), Value::Integer(1));
                fields.insert("decimalplaces".to_string(), Value::Integer(0));
                fields.insert("readonly".to_string(), Value::Boolean(false));
                fields.insert("hexadecimal".to_string(), Value::Boolean(false));
                fields.insert("thousandsseparator".to_string(), Value::Boolean(false));
                fields.insert("textalign".to_string(), Value::String("Left".to_string()));
                fields.insert("updownalign".to_string(), Value::String("Right".to_string()));
                fields.insert("interceptarrowkeys".to_string(), Value::Boolean(true));
            }
            "treeview" => {
                fields.insert("nodes".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("selectednode".to_string(), Value::Nothing);
                fields.insert("pathseparator".to_string(), Value::String("\\".to_string()));
                fields.insert("checkboxes".to_string(), Value::Boolean(false));
                fields.insert("showlines".to_string(), Value::Boolean(true));
                fields.insert("showrootlines".to_string(), Value::Boolean(true));
                fields.insert("showplusminus".to_string(), Value::Boolean(true));
                fields.insert("shownodetoolTips".to_lowercase(), Value::Boolean(false));
                fields.insert("fullrowselect".to_string(), Value::Boolean(false));
                fields.insert("hideselection".to_string(), Value::Boolean(true));
                fields.insert("hottracking".to_string(), Value::Boolean(false));
                fields.insert("labeledit".to_string(), Value::Boolean(false));
                fields.insert("scrollable".to_string(), Value::Boolean(true));
                fields.insert("sorted".to_string(), Value::Boolean(false));
                fields.insert("indent".to_string(), Value::Integer(19));
                fields.insert("itemheight".to_string(), Value::Integer(16));
                fields.insert("imagelist".to_string(), Value::Nothing);
                fields.insert("imageindex".to_string(), Value::Integer(-1));
                fields.insert("selectedimageindex".to_string(), Value::Integer(-1));
                fields.insert("topnode".to_string(), Value::Nothing);
                fields.insert("visiblecount".to_string(), Value::Integer(0));
            }
            "listview" => {
                fields.insert("items".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("columns".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("view".to_string(), Value::Integer(1));
                fields.insert("fullrowselect".to_string(), Value::Boolean(false));
                fields.insert("gridlines".to_string(), Value::Boolean(false));
                fields.insert("checkboxes".to_string(), Value::Boolean(false));
                fields.insert("multiselect".to_string(), Value::Boolean(true));
                fields.insert("selecteditems".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("selectedindices".to_string(), Value::Array(vec![]));
                fields.insert("groups".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("showgroups".to_string(), Value::Boolean(false));
                fields.insert("sorting".to_string(), Value::String("None".to_string()));
                fields.insert("labeledit".to_string(), Value::Boolean(false));
                fields.insert("labelwrap".to_string(), Value::Boolean(true));
                fields.insert("allowcolumnreorder".to_string(), Value::Boolean(false));
                fields.insert("headerStyle".to_lowercase(), Value::String("Clickable".to_string()));
                fields.insert("hideselection".to_string(), Value::Boolean(true));
                fields.insert("hottracking".to_string(), Value::Boolean(false));
                fields.insert("hoveroSelection".to_lowercase(), Value::Boolean(false));
                fields.insert("activation".to_string(), Value::String("Standard".to_string()));
                fields.insert("scrollable".to_string(), Value::Boolean(true));
                fields.insert("showitemtooltips".to_string(), Value::Boolean(false));
                fields.insert("tilelayout".to_string(), Value::Nothing);
                fields.insert("topitem".to_string(), Value::Nothing);
                fields.insert("focuseditem".to_string(), Value::Nothing);
                fields.insert("imagelist".to_string(), Value::Nothing);
                fields.insert("smallimagelist".to_string(), Value::Nothing);
                fields.insert("largeimagelist".to_string(), Value::Nothing);
                fields.insert("stateimagelist".to_string(), Value::Nothing);
            }
            "datagridview" => {
                fields.insert("rows".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("columns".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("selectedrows".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("allowusertoaddrows".to_string(), Value::Boolean(true));
                fields.insert("allowusertodeleterows".to_string(), Value::Boolean(true));
                fields.insert("readonly".to_string(), Value::Boolean(false));
                fields.insert("datasource".to_string(), Value::Nothing);
                fields.insert("datamember".to_string(), Value::String(String::new()));
                fields.insert("autogeneratecolumns".to_string(), Value::Boolean(true));
                fields.insert("multiselect".to_string(), Value::Boolean(true));
                fields.insert("selectionmode".to_string(), Value::String("RowHeaderSelect".to_string()));
                fields.insert("autosizecolumnsmode".to_string(), Value::String("None".to_string()));
                fields.insert("autosizerowsmode".to_string(), Value::String("None".to_string()));
                fields.insert("allowusertoresizecolumns".to_string(), Value::Boolean(true));
                fields.insert("allowusertoresizerows".to_string(), Value::Boolean(true));
                fields.insert("allowusertoordercolumns".to_string(), Value::Boolean(false));
                fields.insert("columnheadersborderstyle".to_string(), Value::String("Raised".to_string()));
                fields.insert("columnheadersvisible".to_string(), Value::Boolean(true));
                fields.insert("rowheadersborderstyle".to_string(), Value::String("Raised".to_string()));
                fields.insert("rowheadersvisible".to_string(), Value::Boolean(true));
                fields.insert("rowheaderswidth".to_string(), Value::Integer(43));
                fields.insert("editmode".to_string(), Value::String("EditOnKeystrokeOrF2".to_string()));
                fields.insert("gridcolor".to_string(), Value::String(String::new()));
                fields.insert("borderstyle".to_string(), Value::String("FixedSingle".to_string()));
                fields.insert("cellborderstyle".to_string(), Value::String("Single".to_string()));
                fields.insert("clipboardcopymode".to_string(), Value::String("EnableWithAutoHeaderText".to_string()));
                fields.insert("scrollbars".to_string(), Value::String("Both".to_string()));
                fields.insert("currentcell".to_string(), Value::Nothing);
                fields.insert("currentrow".to_string(), Value::Nothing);
                fields.insert("selectedcells".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("selectedcolumns".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("firstdisplayedcell".to_string(), Value::Nothing);
                fields.insert("firstdisplayedscrollingrowindex".to_string(), Value::Integer(0));
                fields.insert("firstdisplayedscrollingcolumnindex".to_string(), Value::Integer(0));
                fields.insert("newrowindex".to_string(), Value::Integer(-1));
                fields.insert("sortedcolumn".to_string(), Value::Nothing);
                fields.insert("sortorder".to_string(), Value::String("None".to_string()));
            }
            "tabcontrol" => {
                fields.insert("tabpages".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("selectedindex".to_string(), Value::Integer(0));
                fields.insert("selectedtab".to_string(), Value::Nothing);
                fields.insert("alignment".to_string(), Value::String("Top".to_string()));
                fields.insert("appearance".to_string(), Value::String("Normal".to_string()));
                fields.insert("multiline".to_string(), Value::Boolean(false));
                fields.insert("sizemode".to_string(), Value::String("Normal".to_string()));
                fields.insert("hottrack".to_string(), Value::Boolean(false));
                fields.insert("itemsize".to_string(), Value::Nothing);
                fields.insert("padding".to_string(), Value::Nothing);
                fields.insert("showtooltips".to_string(), Value::Boolean(false));
                fields.insert("imagelist".to_string(), Value::Nothing);
            }
            "tabpage" => {
                fields.insert("text".to_string(), Value::String(String::new()));
                fields.insert("tooltip".to_string(), Value::String(String::new()));
                fields.insert("imageindex".to_string(), Value::Integer(-1));
                fields.insert("imagekey".to_string(), Value::String(String::new()));
                fields.insert("borderstyle".to_string(), Value::String("None".to_string()));
                fields.insert("autosize".to_string(), Value::Boolean(false));
                fields.insert("autoscroll".to_string(), Value::Boolean(false));
                fields.insert("usevisualstylebackcolor".to_string(), Value::Boolean(true));
                fields.insert("padding".to_string(), Value::Integer(3));
                fields.insert("controls".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
            }
            "menustrip" | "contextmenustrip" => {
                fields.insert("items".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("dock".to_string(), Value::Integer(1)); // Top
                fields.insert("rendermode".to_string(), Value::String("ManagerRenderMode".to_string()));
                fields.insert("showitemtooltips".to_string(), Value::Boolean(false));
                fields.insert("grabrendermode".to_string(), Value::String("VisualStyles".to_string()));
                fields.insert("stretch".to_string(), Value::Boolean(true));
                fields.insert("layoutstyle".to_string(), Value::String("HorizontalStackWithOverflow".to_string()));
                fields.insert("imagescalingsize".to_string(), Value::String("16, 16".to_string()));
            }
            "statusstrip" => {
                fields.insert("items".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("dock".to_string(), Value::Integer(2)); // Bottom
                fields.insert("rendermode".to_string(), Value::String("ManagerRenderMode".to_string()));
                fields.insert("showitemtooltips".to_string(), Value::Boolean(false));
                fields.insert("sizinggrip".to_string(), Value::Boolean(true));
                fields.insert("stretch".to_string(), Value::Boolean(true));
                fields.insert("layoutstyle".to_string(), Value::String("Table".to_string()));
            }
            "toolstripstatuslabel" => {
                fields.insert("text".to_string(), Value::String(String::new()));
                fields.insert("spring".to_string(), Value::Boolean(false));
                fields.insert("autosize".to_string(), Value::Boolean(true));
                fields.insert("bordersides".to_string(), Value::String("None".to_string()));
                fields.insert("borderstyle".to_string(), Value::String("None".to_string()));
                fields.insert("islink".to_string(), Value::Boolean(false));
                fields.insert("alignment".to_string(), Value::String("Left".to_string()));
                fields.insert("image".to_string(), Value::Nothing);
                fields.insert("imagealign".to_string(), Value::String("MiddleLeft".to_string()));
                fields.insert("imageindex".to_string(), Value::Integer(-1));
                fields.insert("tooltip".to_string(), Value::String(String::new()));
            }
            "toolstripmenuitem" => {
                fields.insert("dropdownitems".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("checked".to_string(), Value::Boolean(false));
                fields.insert("shortcutkeys".to_string(), Value::Integer(0));
                fields.insert("checkstate".to_string(), Value::String("Unchecked".to_string()));
                fields.insert("checkonclick".to_string(), Value::Boolean(false));
                fields.insert("showshortcutkeys".to_string(), Value::Boolean(true));
                fields.insert("shortcutkeyDisplayString".to_lowercase(), Value::String(String::new()));
                fields.insert("image".to_string(), Value::Nothing);
                fields.insert("imagealign".to_string(), Value::String("MiddleLeft".to_string()));
                fields.insert("imageindex".to_string(), Value::Integer(-1));
                fields.insert("imagescaling".to_string(), Value::String("SizeToFit".to_string()));
                fields.insert("tooltip".to_string(), Value::String(String::new()));
                fields.insert("alignment".to_string(), Value::String("Left".to_string()));
                fields.insert("autosize".to_string(), Value::Boolean(true));
                fields.insert("displaystyle".to_string(), Value::String("ImageAndText".to_string()));
                fields.insert("textalign".to_string(), Value::String("MiddleLeft".to_string()));
            }
            "richtextbox" => {
                fields.insert("readonly".to_string(), Value::Boolean(false));
                fields.insert("scrollbars".to_string(), Value::Integer(3));
                fields.insert("multiline".to_string(), Value::Boolean(true));
                fields.insert("wordwrap".to_string(), Value::Boolean(true));
                fields.insert("maxlength".to_string(), Value::Integer(2147483647));
                fields.insert("rtf".to_string(), Value::String(String::new()));
                fields.insert("selectedtext".to_string(), Value::String(String::new()));
                fields.insert("selectionstart".to_string(), Value::Integer(0));
                fields.insert("selectionlength".to_string(), Value::Integer(0));
                fields.insert("selectioncolor".to_string(), Value::String(String::new()));
                fields.insert("selectionfont".to_string(), Value::Nothing);
                fields.insert("selectionbullet".to_string(), Value::Boolean(false));
                fields.insert("selectionindent".to_string(), Value::Integer(0));
                fields.insert("selectionalignment".to_string(), Value::String("Left".to_string()));
                fields.insert("selectionhangingindent".to_string(), Value::Integer(0));
                fields.insert("selectionbackcolor".to_string(), Value::String(String::new()));
                fields.insert("zoomfactor".to_string(), Value::Double(1.0));
                fields.insert("modified".to_string(), Value::Boolean(false));
                fields.insert("detecturls".to_string(), Value::Boolean(true));
                fields.insert("hideselection".to_string(), Value::Boolean(true));
                fields.insert("acceptstab".to_string(), Value::Boolean(false));
                fields.insert("bulletindent".to_string(), Value::Integer(0));
                fields.insert("autowordselection".to_string(), Value::Boolean(false));
                fields.insert("textlength".to_string(), Value::Integer(0));
                fields.insert("lines".to_string(), Value::Array(vec![]));
            }
            "datetimepicker" => {
                fields.insert("value".to_string(), Value::String(String::new()));
                fields.insert("format".to_string(), Value::String("Long".to_string()));
                fields.insert("customformat".to_string(), Value::String(String::new()));
                fields.insert("mindate".to_string(), Value::String("1/1/1753".to_string()));
                fields.insert("maxdate".to_string(), Value::String("12/31/9998".to_string()));
                fields.insert("showcheckbox".to_string(), Value::Boolean(false));
                fields.insert("checked".to_string(), Value::Boolean(true));
                fields.insert("showupdown".to_string(), Value::Boolean(false));
                fields.insert("calendarforecolor".to_string(), Value::String(String::new()));
                fields.insert("calendarmonthbackground".to_string(), Value::String(String::new()));
                fields.insert("calendartitlebackcolor".to_string(), Value::String(String::new()));
                fields.insert("calendartitleforecolor".to_string(), Value::String(String::new()));
                fields.insert("calendartrailingforecolor".to_string(), Value::String(String::new()));
                fields.insert("righttoLeft".to_lowercase(), Value::String("No".to_string()));
                fields.insert("dropdownalign".to_string(), Value::String("Left".to_string()));
            }
            "linklabel" => {
                fields.insert("linkcolor".to_string(), Value::String("#0066cc".to_string()));
                fields.insert("visitedlinkcolor".to_string(), Value::String("#800080".to_string()));
                fields.insert("activelinkcolor".to_string(), Value::String("Red".to_string()));
                fields.insert("disabledlinkcolor".to_string(), Value::String(String::new()));
                fields.insert("linkvisited".to_string(), Value::Boolean(false));
                fields.insert("linkbehavior".to_string(), Value::String("SystemDefault".to_string()));
                fields.insert("linkarea".to_string(), Value::Nothing);
                fields.insert("links".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("textalign".to_string(), Value::String("TopLeft".to_string()));
                fields.insert("autosize".to_string(), Value::Boolean(true));
            }
            "toolstrip" => {
                fields.insert("items".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("dock".to_string(), Value::Integer(1)); // Top
                fields.insert("grabrendermode".to_string(), Value::String("VisualStyles".to_string()));
                fields.insert("rendermode".to_string(), Value::String("ManagerRenderMode".to_string()));
                fields.insert("stretch".to_string(), Value::Boolean(true));
                fields.insert("showitemtooltips".to_string(), Value::Boolean(true));
                fields.insert("imageScalingSize".to_lowercase(), Value::String("16, 16".to_string()));
                fields.insert("layoutstyle".to_string(), Value::String("HorizontalStackWithOverflow".to_string()));
            }
            "trackbar" => {
                fields.insert("value".to_string(), Value::Integer(0));
                fields.insert("minimum".to_string(), Value::Integer(0));
                fields.insert("maximum".to_string(), Value::Integer(10));
                fields.insert("tickfrequency".to_string(), Value::Integer(1));
                fields.insert("smallchange".to_string(), Value::Integer(1));
                fields.insert("largechange".to_string(), Value::Integer(5));
                fields.insert("orientation".to_string(), Value::String("Horizontal".to_string()));
                fields.insert("tickstyle".to_string(), Value::String("BottomRight".to_string()));
                fields.insert("righttoleft".to_string(), Value::String("No".to_string()));
                fields.insert("righttoleftlayout".to_string(), Value::Boolean(false));
            }
            "maskedtextbox" => {
                fields.insert("mask".to_string(), Value::String(String::new()));
                fields.insert("promptchar".to_string(), Value::String("_".to_string()));
                fields.insert("maskcompleted".to_string(), Value::Boolean(false));
                fields.insert("maskfull".to_string(), Value::Boolean(false));
                fields.insert("readonly".to_string(), Value::Boolean(false));
                fields.insert("hidepromptonleave".to_string(), Value::Boolean(false));
                fields.insert("beepOnError".to_lowercase(), Value::Boolean(false));
                fields.insert("allowpromptasInput".to_lowercase(), Value::Boolean(true));
                fields.insert("asciionly".to_string(), Value::Boolean(false));
                fields.insert("cutcopymaskinclprompt".to_string(), Value::Boolean(false));
                fields.insert("insertkeymode".to_string(), Value::String("Default".to_string()));
                fields.insert("rejectstringoncontroltext".to_string(), Value::Boolean(false));
                fields.insert("resetonprompt".to_string(), Value::Boolean(true));
                fields.insert("resetonspace".to_string(), Value::Boolean(true));
                fields.insert("skipliterals".to_string(), Value::Boolean(true));
                fields.insert("textalign".to_string(), Value::String("Left".to_string()));
                fields.insert("textmaskformat".to_string(), Value::String("IncludeLiterals".to_string()));
                fields.insert("validatingtype".to_string(), Value::Nothing);
            }
            "splitcontainer" => {
                fields.insert("orientation".to_string(), Value::String("Vertical".to_string()));
                fields.insert("splitterdistance".to_string(), Value::Integer(100));
                fields.insert("panel1".to_string(), Value::Nothing);
                fields.insert("panel2".to_string(), Value::Nothing);
                fields.insert("splitterincrement".to_string(), Value::Integer(1));
                fields.insert("splitterwidth".to_string(), Value::Integer(4));
                fields.insert("fixedpanel".to_string(), Value::String("None".to_string()));
                fields.insert("issplitterfixed".to_string(), Value::Boolean(false));
                fields.insert("panel1collapsed".to_string(), Value::Boolean(false));
                fields.insert("panel2collapsed".to_string(), Value::Boolean(false));
                fields.insert("panel1minsize".to_string(), Value::Integer(25));
                fields.insert("panel2minsize".to_string(), Value::Integer(25));
                fields.insert("borderstyle".to_string(), Value::String("None".to_string()));
            }
            "flowlayoutpanel" => {
                fields.insert("flowdirection".to_string(), Value::String("LeftToRight".to_string()));
                fields.insert("wrapcontents".to_string(), Value::Boolean(true));
                fields.insert("autosize".to_string(), Value::Boolean(false));
                fields.insert("autosizemode".to_string(), Value::String("GrowOnly".to_string()));
                fields.insert("autoscroll".to_string(), Value::Boolean(false));
                fields.insert("borderstyle".to_string(), Value::String("None".to_string()));
                fields.insert("padding".to_string(), Value::Integer(0));
            }
            "tablelayoutpanel" => {
                fields.insert("columncount".to_string(), Value::Integer(2));
                fields.insert("rowcount".to_string(), Value::Integer(2));
                fields.insert("autosize".to_string(), Value::Boolean(false));
                fields.insert("autoscroll".to_string(), Value::Boolean(false));
                fields.insert("borderstyle".to_string(), Value::String("None".to_string()));
                fields.insert("cellborderstyle".to_string(), Value::String("None".to_string()));
                fields.insert("columnstyles".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("rowstyles".to_string(), Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                fields.insert("growstyle".to_string(), Value::String("AddRows".to_string()));
                fields.insert("padding".to_string(), Value::Integer(0));
            }
            "monthcalendar" => {
                fields.insert("selectionstart".to_string(), Value::String(String::new()));
                fields.insert("selectionend".to_string(), Value::String(String::new()));
                fields.insert("todaydate".to_string(), Value::String(String::new()));
                fields.insert("maxselectioncount".to_string(), Value::Integer(7));
                fields.insert("showtoday".to_string(), Value::Boolean(true));
                fields.insert("showtodaycircle".to_string(), Value::Boolean(true));
                fields.insert("showweeknumbers".to_string(), Value::Boolean(false));
                fields.insert("firstdayofweek".to_string(), Value::String("Default".to_string()));
                fields.insert("calendarmonth".to_string(), Value::Integer(1));
                fields.insert("calendaryear".to_string(), Value::Integer(2026));
                fields.insert("mindate".to_string(), Value::String(String::new()));
                fields.insert("maxdate".to_string(), Value::String(String::new()));
                fields.insert("selectionrange".to_string(), Value::Nothing);
                fields.insert("bolddates".to_string(), Value::Array(vec![]));
                fields.insert("annuallyboldeddates".to_string(), Value::Array(vec![]));
                fields.insert("monthlyboldeddates".to_string(), Value::Array(vec![]));
                fields.insert("scrollchange".to_string(), Value::Integer(1));
            }
            "hscrollbar" | "vscrollbar" => {
                fields.insert("value".to_string(), Value::Integer(0));
                fields.insert("minimum".to_string(), Value::Integer(0));
                fields.insert("maximum".to_string(), Value::Integer(100));
                fields.insert("smallchange".to_string(), Value::Integer(1));
                fields.insert("largechange".to_string(), Value::Integer(10));
            }
            "tooltip" => {
                fields.insert("active".to_string(), Value::Boolean(true));
                fields.insert("autopopdelay".to_string(), Value::Integer(5000));
                fields.insert("initialdelay".to_string(), Value::Integer(500));
                fields.insert("reshowdelay".to_string(), Value::Integer(100));
                fields.insert("showalways".to_string(), Value::Boolean(false));
                fields.insert("istooltipon".to_string(), Value::Boolean(false));
                fields.insert("usefading".to_string(), Value::Boolean(true));
                fields.insert("useanimation".to_string(), Value::Boolean(true));
                fields.insert("tooltipicon".to_string(), Value::Integer(0));
                fields.insert("tooltiptitle".to_string(), Value::String(String::new()));
                fields.insert("stripampersands".to_string(), Value::Boolean(false));
                fields.insert("__tooltips".to_string(), Value::Nothing);
            }
            _ => {}
        }
    }

    /// Create a System.Windows.Forms.MouseEventArgs object.
    pub fn make_mouse_event_args(button: i32, clicks: i32, x: i32, y: i32, delta: i32) -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("MouseEventArgs".to_string()));
        fields.insert("button".to_string(), Value::Integer(button));
        fields.insert("clicks".to_string(), Value::Integer(clicks));
        fields.insert("x".to_string(), Value::Integer(x));
        fields.insert("y".to_string(), Value::Integer(y));
        fields.insert("delta".to_string(), Value::Integer(delta));
        fields.insert("location".to_string(), Self::make_point(x, y));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "MouseEventArgs".to_string(),
            fields,
        })))
    }

    /// Create a System.Windows.Forms.KeyEventArgs object.
    pub fn make_key_event_args(key_code: i32, shift: bool, ctrl: bool, alt: bool) -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("KeyEventArgs".to_string()));
        fields.insert("keycode".to_string(), Value::Integer(key_code));
        fields.insert("keyvalue".to_string(), Value::Integer(key_code));
        fields.insert("keydata".to_string(), Value::Integer(key_code));
        fields.insert("shift".to_string(), Value::Boolean(shift));
        fields.insert("control".to_string(), Value::Boolean(ctrl));
        fields.insert("alt".to_string(), Value::Boolean(alt));
        fields.insert("handled".to_string(), Value::Boolean(false));
        fields.insert("suppresskeypress".to_string(), Value::Boolean(false));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "KeyEventArgs".to_string(),
            fields,
        })))
    }

    /// Create a System.Windows.Forms.KeyPressEventArgs object.
    pub fn make_key_press_event_args(key_char: char) -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("KeyPressEventArgs".to_string()));
        fields.insert("keychar".to_string(), Value::Char(key_char));
        fields.insert("handled".to_string(), Value::Boolean(false));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "KeyPressEventArgs".to_string(),
            fields,
        })))
    }

    /// Create a System.Windows.Forms.FormClosingEventArgs object.
    pub fn make_form_closing_event_args(close_reason: i32) -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("FormClosingEventArgs".to_string()));
        fields.insert("cancel".to_string(), Value::Boolean(false));
        // CloseReason enum: 0=None, 1=WindowsShutDown, 2=MdiFormClosing, 3=UserClosing, 4=TaskManagerClosing, 5=FormOwnerClosing, 6=ApplicationExitCall
        fields.insert("closereason".to_string(), Value::Integer(close_reason));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "FormClosingEventArgs".to_string(),
            fields,
        })))
    }

    /// Create a System.Windows.Forms.FormClosedEventArgs object.
    pub fn make_form_closed_event_args(close_reason: i32) -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("FormClosedEventArgs".to_string()));
        fields.insert("closereason".to_string(), Value::Integer(close_reason));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "FormClosedEventArgs".to_string(),
            fields,
        })))
    }

    /// Create a System.Windows.Forms.PaintEventArgs stub.
    pub fn make_paint_event_args() -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("PaintEventArgs".to_string()));
        // Graphics and ClipRectangle are stubs
        fields.insert("cliprectangle".to_string(), Value::String("0, 0, 0, 0".to_string()));
        Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
            class_name: "PaintEventArgs".to_string(),
            fields,
        })))
    }

    /// Build standard .NET event args: returns (sender, e) pair for a given event name.
    /// sender is the control object (or Nothing if not found), e is the appropriate EventArgs subclass.
    /// `event_data` provides real values from the UI framework when available.
    pub fn make_event_handler_args(&self, control_name: &str, event_name: &str) -> Vec<Value> {
        self.make_event_handler_args_with_data(control_name, event_name, None)
    }

    /// Build event args with optional concrete event data from the UI layer.
    pub fn make_event_handler_args_with_data(&self, control_name: &str, event_name: &str, data: Option<&EventData>) -> Vec<Value> {
        let sender = self.resolve_control_as_sender(control_name);

        let event_lower = event_name.to_lowercase();
        let e = match event_lower.as_str() {
            "mouseclick" | "mousedown" | "mouseup" | "mousemove" | "mousedoubleclick" | "mouseenter" | "mouseleave" | "mousewheel" => {
                if let Some(EventData::Mouse { button, clicks, x, y, delta }) = data {
                    Self::make_mouse_event_args(*button, *clicks, *x, *y, *delta)
                } else {
                    Self::make_mouse_event_args(0, 1, 0, 0, 0)
                }
            }
            "click" | "doubleclick" | "dblclick" => {
                // Click uses EventArgs but if we have mouse data, embed it anyway
                if let Some(EventData::Mouse { button, clicks, x, y, delta }) = data {
                    Self::make_mouse_event_args(*button, *clicks, *x, *y, *delta)
                } else {
                    Self::make_event_args()
                }
            }
            "keydown" | "keyup" => {
                if let Some(EventData::Key { key_code, shift, ctrl, alt }) = data {
                    Self::make_key_event_args(*key_code, *shift, *ctrl, *alt)
                } else {
                    Self::make_key_event_args(0, false, false, false)
                }
            }
            "keypress" => {
                if let Some(EventData::KeyPress { key_char }) = data {
                    Self::make_key_press_event_args(*key_char)
                } else {
                    Self::make_key_press_event_args('\0')
                }
            }
            "formclosing" =>
                Self::make_form_closing_event_args(3), // UserClosing
            "formclosed" =>
                Self::make_form_closed_event_args(3),
            "paint" =>
                Self::make_paint_event_args(),
            // All other events use base EventArgs
            _ => Self::make_event_args(),
        };

        vec![sender, e]
    }

    /// Resolve a control name to a Value suitable for `sender` in event handler args.
    fn resolve_control_as_sender(&self, control_name: &str) -> Value {
        // Try looking up "Me.ControlName" field on the form instance
        // or "control_name" as a standalone variable
        if let Ok(val) = self.env.get(control_name) {
            if matches!(val, Value::Object(_)) {
                return val;
            }
        }
        // Try with the form instance field
        if let Some(obj) = &self.current_object {
            let borrowed = obj.borrow();
            let key = control_name.to_lowercase();
            if let Some(val) = borrowed.fields.get(&key) {
                return val.clone();
            }
        }
        // Return Nothing if we can't find the control
        Value::Nothing
    }

    /// Legacy: register resources as simple string map (backward compat)
    pub fn register_resources(&mut self, resources: HashMap<String, String>) {
        self.resources = resources.clone();
        let entries: Vec<crate::ResourceEntry> = resources.into_iter()
            .map(|(k, v)| crate::ResourceEntry::string(k, v))
            .collect();
        self.register_resource_entries(entries);
    }

    /// Register typed resource entries – creates the My.Resources namespace with proper
    /// typed values and a ResourceManager sub-object with GetString()/GetObject() methods.
    pub fn register_resource_entries(&mut self, entries: Vec<crate::ResourceEntry>) {
        // Keep a flat string map for backward compat
        for e in &entries {
            self.resources.insert(e.name.clone(), e.value.clone());
        }
        self.resource_entries = entries.clone();

        // Build Resource fields: each resource is a field on My.Resources
        let mut res_fields = HashMap::new();
        for entry in &entries {
            let key = entry.name.to_lowercase();
            let val = match entry.resource_type.as_str() {
                "image" | "icon" | "audio" | "file" => {
                    // File-based resource: create an object with path/type metadata
                    let mut fields = HashMap::new();
                    fields.insert("__type".to_string(), Value::String(format!("Resource.{}", entry.resource_type)));
                    fields.insert("name".to_string(), Value::String(entry.name.clone()));
                    fields.insert("filepath".to_string(), Value::String(entry.file_path.clone().unwrap_or_default()));
                    fields.insert("resourcetype".to_string(), Value::String(entry.resource_type.clone()));
                    // ToString returns the file path
                    fields.insert("__tostring".to_string(), Value::String(entry.file_path.clone().unwrap_or(entry.value.clone())));
                    Value::Object(Rc::new(RefCell::new(ObjectData {
                        class_name: format!("Resource.{}", entry.resource_type),
                        fields,
                    })))
                }
                _ => Value::String(entry.value.clone()),
            };
            // Store both lowercase and original-case access
            res_fields.insert(key, val.clone());
            res_fields.insert(entry.name.clone(), val);
        }

        // Build the ResourceManager sub-object
        // It holds all entries and supports GetString(key) / GetObject(key)
        let mut rm_fields = HashMap::new();
        rm_fields.insert("__type".to_string(), Value::String("ResourceManager".to_string()));
        // Store all entries as a lookup map within the ResourceManager object
        for entry in &entries {
            rm_fields.insert(
                format!("__res_{}", entry.name.to_lowercase()),
                Value::String(entry.value.clone()),
            );
        }
        let rm_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "ResourceManager".to_string(),
            fields: rm_fields,
        })));
        res_fields.insert("resourcemanager".to_string(), rm_obj);

        let res_obj_data = ObjectData {
            class_name: "My.Resources".to_string(),
            fields: res_fields,
        };
        let res_obj = Value::Object(Rc::new(RefCell::new(res_obj_data)));

        // Create or Update My namespace
        if let Ok(my_val) = self.env.get("my") {
             if let Value::Object(my_obj) = my_val {
                 my_obj.borrow_mut().fields.insert("resources".to_string(), res_obj);
             }
        } else {
             let mut my_fields = HashMap::new();
             my_fields.insert("resources".to_string(), res_obj);
             
             let my_obj_data = ObjectData {
                 class_name: "My".to_string(),
                 fields: my_fields,
             };
             let my_obj = Value::Object(Rc::new(RefCell::new(my_obj_data)));
             self.env.define("my", my_obj);
        }
    }

    // ── Network handle management ──────────────────────────────────────
    fn alloc_net_handle(&mut self, handle: crate::builtins::networking::NetHandle) -> i64 {
        let id = self.next_net_handle_id;
        self.next_net_handle_id += 1;
        self.net_handles.insert(id, handle);
        id
    }

    fn remove_net_handle(&mut self, id: i64) {
        self.net_handles.remove(&id);
    }

    fn register_builtin_constants(&mut self) {
        self.env.define_const("vbcrlf", Value::String("\r\n".to_string()));
        self.env.define_const("vbnewline", Value::String("\r\n".to_string()));
        self.env.define_const("vbtab", Value::String("\t".to_string()));
        self.env.define_const("vblf", Value::String("\n".to_string()));
        self.env.define_const("vbcr", Value::String("\r".to_string()));
        self.env.define_const("vbnullchar", Value::String("\0".to_string()));
        self.env.define_const("vbnullstring", Value::String(String::new()));
        self.env.define_const("vbback", Value::String("\x08".to_string()));
        self.env.define_const("vbformfeed", Value::String("\x0C".to_string()));
        self.env.define_const("vbbinarycompare", Value::Integer(0));
        self.env.define_const("vbtextcompare", Value::Integer(1));

        // Math constants
        self.env.define_const("math.pi", Value::Double(std::f64::consts::PI));
        self.env.define_const("math.e", Value::Double(std::f64::consts::E));

        // Environment.NewLine
        if cfg!(target_os = "windows") {
            self.env.define_const("environment.newline", Value::String("\r\n".to_string()));
        } else {
            self.env.define_const("environment.newline", Value::String("\n".to_string()));
        }

        // RegexOptions constants
        self.env.define_const("regexoptions.none", Value::Integer(0));
        self.env.define_const("regexoptions.ignorecase", Value::Integer(1));
        self.env.define_const("regexoptions.multiline", Value::Integer(2));
        self.env.define_const("regexoptions.singleline", Value::Integer(16));
        self.env.define_const("regexoptions.ignorepatternwhitespace", Value::Integer(32));

        // StringSplitOptions constants
        self.env.define_const("stringsplitoptions.none", Value::Integer(0));
        self.env.define_const("stringsplitoptions.removeemptyentries", Value::Integer(1));
        self.env.define_const("stringsplitoptions.trimentriesandremoveemptyentries", Value::Integer(3));

        // StringComparison constants
        self.env.define_const("stringcomparison.ordinal", Value::Integer(4));
        self.env.define_const("stringcomparison.ordinalignorecase", Value::Integer(5));
        self.env.define_const("stringcomparison.currentculture", Value::Integer(0));
        self.env.define_const("stringcomparison.currentcultureignorecase", Value::Integer(1));
    }

    pub fn run(&mut self, program: &Program) -> Result<(), RuntimeError> {
        // First pass: collect all declarations
        for decl in &program.declarations {
            self.declare(decl)?;
        }

        // Second pass: execute statements
        for stmt in &program.statements {
            self.execute(stmt)?;
        }

        Ok(())
    }

    /// Set the command-line arguments available to the VB program.
    /// These are the arguments *after* the project file on the vybe command line.
    /// We prepend the program name (like .NET's Environment.GetCommandLineArgs).
    pub fn set_command_line_args(&mut self, args: Vec<String>) {
        let mut full = vec!["vybe".to_string()];
        full.extend(args);
        self.command_line_args = full;
    }

    pub fn load_module(&mut self, module_name: &str, program: &Program) -> Result<(), RuntimeError> {
        // Define module as a namespace (string proxy) so Module.Member works
        self.env.define(module_name, Value::String(module_name.to_lowercase()));

        // Set current module for scoping
        let prev_module = self.current_module.clone();
        self.current_module = Some(module_name.to_string());

        // Load declarations (will be prefixed with module name)
        for decl in &program.declarations {
            self.declare(decl)?;
        }

        // Execute module-level statements if any
        for stmt in &program.statements {
            self.execute(stmt)?;
        }

        // Restore previous module
        self.current_module = prev_module;
        Ok(())
    }

    /// Load a VB.NET code file into **global** scope.
    ///
    /// Unlike `load_module`, this does NOT prefix declarations with a module
    /// name.  The parser already flattens `Module X ... End Module` contents
    /// into top-level declarations, so everything defined in a Module is
    /// globally accessible — exactly how VB.NET works.
    ///
    /// Use this for `.vbproj` code files.  Use `load_module` when you need a
    /// named scope (VB6 `.bas` files, form code-behind, synthetic helpers).
    pub fn load_code_file(&mut self, program: &Program) -> Result<(), RuntimeError> {
        // No current_module ⇒ declare() registers names without a prefix.
        let prev_module = self.current_module.take();

        for decl in &program.declarations {
            self.declare(decl)?;
        }

        for stmt in &program.statements {
            self.execute(stmt)?;
        }

        self.current_module = prev_module;
        Ok(())
    }

    fn declare(&mut self, decl: &Declaration) -> Result<(), RuntimeError> {
        match decl {
            Declaration::Variable(var) => {
                if !self.env.exists_in_current_scope(var.name.as_str()) {
                    let val = if let Some(expr) = &var.initializer {
                        self.evaluate_expr(expr)?
                    } else {
                        default_value_for_type(&var.name.as_str(), &var.var_type)
                    };
                    
                    let key = if let Some(module) = &self.current_module {
                        format!("{}.{}", module, var.name.as_str()).to_lowercase()
                    } else {
                        var.name.as_str().to_lowercase()
                    };
                    self.env.define(&key, val);
                }
                Ok(())
            }
            Declaration::Constant(const_decl) => {
                // Evaluate the constant expression and define it
                let val = self.evaluate_expr(&const_decl.value)?;
                self.env.define_const(const_decl.name.as_str(), val);
                Ok(())
            }
            Declaration::Sub(sub_decl) => {
                let key = if let Some(module) = &self.current_module {
                    format!("{}.{}", module.to_lowercase(), sub_decl.name.as_str().to_lowercase())
                } else {
                    sub_decl.name.as_str().to_lowercase()
                };
                self.subs.insert(key, sub_decl.clone());
                Ok(())
            }
            Declaration::Function(func_decl) => {
                let key = if let Some(module) = &self.current_module {
                    format!("{}.{}", module.to_lowercase(), func_decl.name.as_str().to_lowercase())
                } else {
                    func_decl.name.as_str().to_lowercase()
                };
                self.functions.insert(key, func_decl.clone());
                Ok(())
            }
            Declaration::Class(class_decl) => {
                let key = if let Some(module) = &self.current_module {
                    format!("{}.{}", module.to_lowercase(), class_decl.name.as_str().to_lowercase())
                } else {
                    class_decl.name.as_str().to_lowercase()
                };

                if let Some(existing) = self.classes.get_mut(&key) {
                    // Merge if either side is partial (VB.NET designer + user code)
                    if existing.is_partial || class_decl.is_partial {
                        existing.fields.extend(class_decl.fields.clone());
                        existing.methods.extend(class_decl.methods.clone());
                        existing.properties.extend(class_decl.properties.clone());
                        if existing.inherits.is_none() {
                            existing.inherits = class_decl.inherits.clone();
                        }
                        if existing.implements.is_empty() && !class_decl.implements.is_empty() {
                            existing.implements = class_decl.implements.clone();
                        }
                        existing.is_partial = existing.is_partial || class_decl.is_partial;
                        existing.is_must_inherit = existing.is_must_inherit || class_decl.is_must_inherit;
                        existing.is_not_inheritable = existing.is_not_inheritable || class_decl.is_not_inheritable;
                    } else {
                        // Replace only when both are non-partial
                        *existing = class_decl.clone();
                    }
                } else {
                    self.classes.insert(key.clone(), class_decl.clone());
                }

                // Enforce NotInheritable
                if let Some(vybe_parser::VBType::Custom(parent_name)) = &class_decl.inherits {
                    let parent_key = self.resolve_class_key(parent_name);
                    if let Some(parent_cls) = parent_key.as_ref().and_then(|k| self.classes.get(k)) {
                        if parent_cls.is_not_inheritable {
                            return Err(RuntimeError::Custom(format!(
                                "Class '{}' cannot inherit from NotInheritable class '{}'",
                                class_decl.name.as_str(),
                                parent_cls.name.as_str()
                            )));
                        }
                    }
                }

                // Register class methods as subs so they can be called by event system
                for method in &class_decl.methods {
                    match method {
                        vybe_parser::ast::MethodDecl::Sub(sub_decl) => {
                            let sub_key = format!("{}.{}", key, sub_decl.name.as_str().to_lowercase());
                            self.subs.insert(sub_key, sub_decl.clone());
                        }
                        vybe_parser::ast::MethodDecl::Function(func_decl) => {
                            let func_key = format!("{}.{}", key, func_decl.name.as_str().to_lowercase());
                            self.functions.insert(func_key, func_decl.clone());
                        }
                    }
                }

                Ok(())
            }
            Declaration::Enum(enum_decl) => {
                // Register each enum member as a constant
                let mut auto_value = 0i32;
                for member in &enum_decl.members {
                    let val = if let Some(expr) = &member.value {
                        let v = self.evaluate_expr(expr)?;
                        auto_value = v.as_integer()? + 1;
                        v
                    } else {
                        let v = Value::Integer(auto_value);
                        auto_value += 1;
                        v
                    };
                    self.env.define_const(member.name.as_str(), val);
                }
                Ok(())
            }
            Declaration::Namespace(ns_decl) => {
                self.declare_namespace(&ns_decl.name, &ns_decl.declarations)
            }
            Declaration::Imports(imp_decl) => {
                self.imports.push(ImportEntry {
                    path: imp_decl.path.clone(),
                    alias: imp_decl.alias.clone(),
                });
                // If the import has an alias, register it as an env variable
                // so `alias.Member` works like a member access.
                if let Some(alias) = &imp_decl.alias {
                    // Try to find an existing namespace object for the path
                    let path_lower = imp_decl.path.to_lowercase();
                    if let Ok(val) = self.env.get(&path_lower) {
                        self.env.define(&alias.to_lowercase(), val);
                    }
                }
                Ok(())
            }
            Declaration::Interface(iface_decl) => {
                let key = if let Some(module) = &self.current_module {
                    format!("{}.{}", module.to_lowercase(), iface_decl.name.as_str().to_lowercase())
                } else {
                    iface_decl.name.as_str().to_lowercase()
                };
                self.interfaces.insert(key, iface_decl.clone());
                Ok(())
            }
            Declaration::Structure(struct_decl) => {
                let key = if let Some(module) = &self.current_module {
                    format!("{}.{}", module.to_lowercase(), struct_decl.name.as_str().to_lowercase())
                } else {
                    struct_decl.name.as_str().to_lowercase()
                };
                // Register structure like a class (value type semantics)
                let class = vybe_parser::ClassDecl {
                    visibility: struct_decl.visibility.clone(),
                    name: struct_decl.name.clone(),
                    is_partial: false,
                    inherits: None,
                    implements: struct_decl.implements.clone(),
                    properties: struct_decl.properties.clone(),
                    methods: struct_decl.methods.clone(),
                    fields: struct_decl.fields.clone(),
                    is_must_inherit: false,
                    is_not_inheritable: true, // structs can't be inherited
                };
                self.classes.insert(key.clone(), class);
                self.structures.insert(key, struct_decl.clone());
                Ok(())
            }
            Declaration::Delegate(del_decl) => {
                let key = if let Some(module) = &self.current_module {
                    format!("{}.{}", module.to_lowercase(), del_decl.name.as_str().to_lowercase())
                } else {
                    del_decl.name.as_str().to_lowercase()
                };
                self.delegates.insert(key, del_decl.clone());
                Ok(())
            }
            Declaration::Event(event_decl) => {
                // Events are registered in the class context — for now just acknowledge
                // The event system already handles AddHandler/RaiseEvent at runtime.
                let _name = event_decl.name.as_str();
                Ok(())
            }
        }
    }

    /// Register all declarations inside a Namespace block, prefixing them
    /// with the fully-qualified namespace name.
    fn declare_namespace(&mut self, ns_name: &str, declarations: &[vybe_parser::Declaration]) -> Result<(), RuntimeError> {
        let prev_module = self.current_module.clone();
        // Set current_module to the namespace so nested declare() calls prefix correctly
        let full_ns = if let Some(ref outer) = prev_module {
            format!("{}.{}", outer, ns_name)
        } else {
            ns_name.to_string()
        };
        self.current_module = Some(full_ns.clone());

        for decl in declarations {
            match decl {
                Declaration::Namespace(inner_ns) => {
                    // Nested namespace: append to current
                    self.declare_namespace(&inner_ns.name, &inner_ns.declarations)?;
                }
                Declaration::Class(class_decl) => {
                    // Register with fully-qualified key AND short name
                    let short_key = class_decl.name.as_str().to_lowercase();
                    let qualified_key = format!("{}.{}", full_ns.to_lowercase(), short_key);

                    // Store in classes under the qualified key
                    if let Some(existing) = self.classes.get_mut(&qualified_key) {
                        if existing.is_partial || class_decl.is_partial {
                            existing.fields.extend(class_decl.fields.clone());
                            existing.methods.extend(class_decl.methods.clone());
                            existing.properties.extend(class_decl.properties.clone());
                            if existing.inherits.is_none() {
                                existing.inherits = class_decl.inherits.clone();
                            }
                            if existing.implements.is_empty() && !class_decl.implements.is_empty() {
                                existing.implements = class_decl.implements.clone();
                            }
                        } else {
                            *existing = class_decl.clone();
                        }
                    } else {
                        self.classes.insert(qualified_key.clone(), class_decl.clone());
                    }

                    // Also register under the short name if not already taken (convenience)
                    if !self.classes.contains_key(&short_key) {
                        self.classes.insert(short_key.clone(), class_decl.clone());
                    }

                    // Record the FQ → key mapping
                    self.namespace_map.insert(qualified_key.clone(), qualified_key.clone());
                    self.namespace_map.insert(short_key.clone(), qualified_key.clone());

                    // Register class methods
                    for method in &class_decl.methods {
                        match method {
                            vybe_parser::ast::MethodDecl::Sub(sub_decl) => {
                                let sub_key = format!("{}.{}", qualified_key, sub_decl.name.as_str().to_lowercase());
                                self.subs.insert(sub_key, sub_decl.clone());
                            }
                            vybe_parser::ast::MethodDecl::Function(func_decl) => {
                                let func_key = format!("{}.{}", qualified_key, func_decl.name.as_str().to_lowercase());
                                self.functions.insert(func_key, func_decl.clone());
                            }
                        }
                    }
                }
                other => {
                    // Subs, Functions, Variables, Constants, Enums — use normal declare()
                    self.declare(other)?;
                }
            }
        }

        self.current_module = prev_module;
        Ok(())
    }

    /// Resolve a class name to the key actually stored in `self.classes`.
    ///
    /// Tries, in order:
    ///   1. exact match (already lowercased)
    ///   2. current-module-qualified  (`module.classname`)
    ///   3. namespace_map lookup (fully-qualified from namespace declarations)
    ///   4. imports-qualified (`imported_ns.classname`)
    ///   5. any key whose last segment matches (`*.classname`)
    fn resolve_class_key(&self, class_name: &str) -> Option<String> {
        let lower = class_name.to_lowercase();

        // 1. Exact match
        if self.classes.contains_key(&lower) {
            return Some(lower);
        }

        // 2. Try module/namespace-qualified
        if let Some(module) = &self.current_module {
            let qualified = format!("{}.{}", module.to_lowercase(), lower);
            if self.classes.contains_key(&qualified) {
                return Some(qualified);
            }
        }

        // 3. Check namespace_map
        if let Some(fq_key) = self.namespace_map.get(&lower) {
            if self.classes.contains_key(fq_key) {
                return Some(fq_key.clone());
            }
        }

        // 4. Try each imported namespace as prefix
        for imp in &self.imports {
            if imp.alias.is_none() {
                let qualified = format!("{}.{}", imp.path.to_lowercase(), lower);
                if self.classes.contains_key(&qualified) {
                    return Some(qualified);
                }
            }
        }

        // 5. Fallback: search for any key ending with `.classname`
        let suffix = format!(".{}", lower);
        self.classes.keys()
            .find(|k| k.ends_with(&suffix))
            .cloned()
    }

    // Helper to collect all fields including inherited ones
    fn collect_fields(&mut self, class_name: &str) -> HashMap<String, Value> {
        let mut fields = HashMap::new();
        
        let resolved = self.resolve_class_key(class_name);
        // 1. Get base class fields first (if any)
        if let Some(cls) = resolved.and_then(|k| self.classes.get(&k).cloned()) {
             if let Some(parent_type) = &cls.inherits {
                 // Resolve parent type to string
                 // VBType::Custom(name) usually
                 let parent_name = match parent_type {
                     vybe_parser::VBType::Custom(n) => Some(n.clone()),
                     vybe_parser::VBType::Object => None, // Object has no fields
                     _ => None, // Primitives don't have fields
                 };
                 
                 if let Some(p_name) = parent_name {
                     let parent_fields = self.collect_fields(&p_name);
                     fields.extend(parent_fields);
                 }
             }
             
             // 2. Add/Override with current class fields
            for field in &cls.fields {
                 let init_val = if let Some(expr) = &field.initializer {
                     self.evaluate_expr(expr).unwrap_or(Value::Nothing) 
                 } else {
                     match &field.var_type {
                         Some(t) => match t {
                             vybe_parser::VBType::Integer => Value::Integer(0),
                             vybe_parser::VBType::Long => Value::Long(0),
                             vybe_parser::VBType::Single => Value::Single(0.0),
                             vybe_parser::VBType::Double => Value::Double(0.0),
                             vybe_parser::VBType::String => Value::String("".to_string()),
                             vybe_parser::VBType::Boolean => Value::Boolean(false),
                             vybe_parser::VBType::Custom(s) => {
                                 let s_lower = s.to_lowercase();
                                 if s_lower.contains("system.windows.forms.") || 
                                    s_lower == "button" || 
                                    s_lower == "label" || 
                                    s_lower == "textbox" || 
                                    s_lower == "checkbox" || 
                                    s_lower == "radiobutton" ||
                                    s_lower == "groupbox" ||
                                    s_lower == "panel" ||
                                    s_lower == "form" ||
                                    s_lower == "datagridview" ||
                                    s_lower == "combobox" ||
                                    s_lower == "listbox" ||
                                    s_lower == "picturebox" ||
                                    s_lower == "timer" ||
                                    s_lower == "toolstrip" ||
                                    s_lower == "menustrip" ||
                                    s_lower == "statusstrip" ||
                                    s_lower == "tabcontrol" ||
                                    s_lower == "richtextbox" ||
                                    s_lower == "progressbar" ||
                                    s_lower == "trackbar" ||
                                    s_lower == "numericupdown" ||
                                    s_lower == "datetimepicker" ||
                                    s_lower == "monthcalendar" ||
                                    s_lower == "treeview" ||
                                    s_lower == "listview" ||
                                    s_lower == "webbrowser" ||
                                    s_lower == "errorprovider" ||
                                    s_lower == "tooltip" ||
                                    s_lower == "bindingnavigator" ||
                                    s_lower == "bindingsource" ||
                                    s_lower == "component" ||
                                    s_lower == "container" ||
                                    s_lower == "contextmenustrip" ||
                                    s_lower == "flowlayoutpanel" ||
                                    s_lower == "tablelayoutpanel" ||
                                    s_lower == "splitcontainer" ||
                                    s_lower == "propertygrid" ||
                                    s_lower == "domainupdown" ||
                                    s_lower == "maskedtextbox" ||
                                    s_lower == "linklabel" ||
                                    s_lower == "printdocument" ||
                                    s_lower == "printpreviewcontrol" ||
                                    s_lower == "printpreviewdialog" ||
                                    s_lower == "pagesetupdialog" ||
                                    s_lower == "printdialog" ||
                                    s_lower == "colordialog" ||
                                    s_lower == "fontdialog" ||
                                    s_lower == "folderbrowserdialog" ||
                                    s_lower == "openfiledialog" ||
                                    s_lower == "savefiledialog" ||
                                    s_lower == "checkedlistbox" ||
                                    s_lower == "splitter" ||
                                    s_lower == "datagrid" ||
                                    s_lower == "usercontrol" ||
                                    s_lower == "toolstripseparator" ||
                                    s_lower == "toolstripbutton" ||
                                    s_lower == "toolstriplabel" ||
                                    s_lower == "toolstripcombobox" ||
                                    s_lower == "toolstripdropdownbutton" ||
                                    s_lower == "toolstripsplitbutton" ||
                                    s_lower == "toolstriptextbox" ||
                                    s_lower == "toolstripprogressbar" ||
                                    s_lower == "helpprovider" ||
                                    s_lower == "dataview"
                                 {
                                     Value::String(field.name.as_str().to_string())
                                 } else {
                                     Value::Nothing
                                 }
                             }
                             _ => Value::Nothing,
                         },
                         None => Value::Nothing,
                     }
                 };
                 fields.insert(field.name.as_str().to_lowercase(), init_val);
             }
        }
        
        fields
    }

    // Helper to find a method in class hierarchy
    fn find_method(&self, class_name: &str, method_name: &str) -> Option<vybe_parser::ast::decl::MethodDecl> {
        let key = self.resolve_class_key(class_name)?;
        if let Some(cls) = self.classes.get(&key) {
            // Check current class
            for method in &cls.methods {
                let m_name = match method {
                    vybe_parser::ast::decl::MethodDecl::Sub(s) => &s.name,
                    vybe_parser::ast::decl::MethodDecl::Function(f) => &f.name,
                };
                if m_name.as_str().eq_ignore_ascii_case(method_name) {
                    return Some(method.clone());
                }
            }
            
            // Check base class
             if let Some(parent_type) = &cls.inherits {
                 let parent_name = match parent_type {
                     vybe_parser::VBType::Custom(n) => Some(n.clone()),
                     _ => None,
                 };
                 if let Some(p_name) = parent_name {
                     return self.find_method(&p_name, method_name);
                 }
             }
        }
        None
    }

    // Helper to find a property in class hierarchy
    fn find_property(&self, class_name: &str, prop_name: &str) -> Option<vybe_parser::ast::decl::PropertyDecl> {
        let key = self.resolve_class_key(class_name).unwrap_or_else(|| class_name.to_lowercase());
        if let Some(cls) = self.classes.get(&key) {
             // Check current class
            for prop in &cls.properties {
                if prop.name.as_str().eq_ignore_ascii_case(prop_name) {
                    return Some(prop.clone());
                }
            }
            
            // Check base class
             if let Some(parent_type) = &cls.inherits {
                 let parent_name = match parent_type {
                     vybe_parser::VBType::Custom(n) => Some(n.clone()),
                     _ => None,
                 };
                 if let Some(p_name) = parent_name {
                     return self.find_property(&p_name, prop_name);
                 }
             }
        }
        None
    }

    /// Get the parent class name from a class's `Inherits` clause.
    pub fn get_parent_class_name(&self, class_name: &str) -> Option<String> {
        let key = self.resolve_class_key(class_name)?;
        let cls = self.classes.get(&key)?;
        match &cls.inherits {
            Some(vybe_parser::VBType::Custom(n)) => Some(n.clone()),
            _ => None,
        }
    }

    /// Find a method starting from the **parent** of the given class.
    /// Used for `MyBase.Method()` — skips the derived class entirely.
    pub fn find_method_in_base(&self, class_name: &str, method_name: &str) -> Option<vybe_parser::ast::decl::MethodDecl> {
        if let Some(parent) = self.get_parent_class_name(class_name) {
            self.find_method(&parent, method_name)
        } else {
            None
        }
    }

    /// Walk the inheritance chain to check if `class_name` is or inherits from `target`.
    /// Used for `TypeOf x Is BaseClass`.
    pub fn is_type_or_base(&self, class_name: &str, target: &str) -> bool {
        if class_name.eq_ignore_ascii_case(target) {
            return true;
        }
        // Also match against unqualified target (e.g. "Form" matches "System.Windows.Forms.Form")
        let class_lower = class_name.to_lowercase();
        let target_lower = target.to_lowercase();
        if class_lower.ends_with(&format!(".{}", target_lower)) || target_lower.ends_with(&format!(".{}", class_lower)) {
            return true;
        }
        // Walk up the hierarchy
        if let Some(parent) = self.get_parent_class_name(class_name) {
            return self.is_type_or_base(&parent, target);
        }
        false
    }


    pub fn execute(&mut self, stmt: &Statement) -> Result<(), RuntimeError> {
        match stmt {
            Statement::Dim(decl) => {
                if let Some(bounds) = &decl.array_bounds {
                    if bounds.len() == 1 {
                        // 1-D array: Dim arr(10) As Integer
                        let size = (self.evaluate_expr(&bounds[0])?.as_integer()? + 1) as usize;
                        let default_val = default_value_for_type("", &decl.var_type);
                        let arr = Value::Array(vec![default_val; size]);
                        self.env.define(decl.name.as_str(), arr);
                    } else {
                        // Multi-dimensional: Dim arr(3, 4) → nested arrays
                        let dims: Vec<usize> = bounds.iter()
                            .map(|b| self.evaluate_expr(b).and_then(|v| v.as_integer()).map(|i| (i + 1) as usize))
                            .collect::<Result<Vec<_>, _>>()?;
                        let default_val = default_value_for_type("", &decl.var_type);
                        let arr = create_multi_dim_array(&dims, &default_val);
                        self.env.define(decl.name.as_str(), arr);
                    }
                } else if let Some(init) = &decl.initializer {
                    // Variable with initializer: Dim x As Integer = 10 or Dim arr() = {1,2,3}
                    let val = self.evaluate_expr(init)?;
                    self.env.define(decl.name.as_str(), val);
                } else {
                    // Regular variable: Dim x As Integer
                    let val = default_value_for_type(&decl.name.as_str(), &decl.var_type);
                    self.env.define(decl.name.as_str(), val);
                }
                Ok(())
            }
            Statement::Const(decl) => {
                // Evaluate the constant expression
                let val = self.evaluate_expr(&decl.value)?;
                // Store as a constant (read-only)
                self.env.define_const(decl.name.as_str(), val);
                Ok(())
            }
            Statement::Assignment { target, value } => {
                let val = self.evaluate_expr(value)?;
                let target_str = target.as_str();
                let target_lower = target_str.to_lowercase();
                
                // 1. Check if it's a local variable (exists in any scope except the global one)
                if self.env.has_local(target_str) {
                    self.env.set(target_str, val)?;
                    return Ok(());
                }

                // 2. Check if it's a field in the current object
                if let Some(obj_rc) = &self.current_object {
                    let mut obj = obj_rc.borrow_mut();
                    // Create or update instance field
                    obj.fields.insert(target_lower, val);
                    return Ok(());
                }

                // 3. Fallback: set in environment (will check global scope or define as global)
                self.env.set(target_str, val)?;
                Ok(())
            }

            Statement::SetAssignment { target, value } => {
                let val = self.evaluate_expr(value)?;
                self.env.set(target.as_str(), val)?;
                Ok(())
            }

            Statement::SyncLock { lock_object, body } => {
                let _lock = self.evaluate_expr(lock_object)?;
                // Verify reference type if strict, but for now just execute body
                // In a single-threaded interpreter, the lock is implicit
                for stmt in body {
                    self.execute(stmt)?;
                }
                Ok(())
            }

            Statement::MemberAssignment { object, member, value } => {
                let val = self.evaluate_expr(value)?;
                let obj_val = self.evaluate_expr(object)?;
                let prop_name = member.as_str().to_string();

                match obj_val {
                    Value::Object(obj_ref) => {
                        let class_key = obj_ref.borrow().class_name.to_lowercase();
                        if let Some(class_decl) = self.classes.get(&class_key).cloned() {
                            for prop in &class_decl.properties {
                                if prop.name.as_str().eq_ignore_ascii_case(&prop_name) {
                                    if let Some((param, body)) = &prop.setter {
                                        let sub = SubDecl {
                                            visibility: prop.visibility,
                                            name: prop.name.clone(),
                                            parameters: vec![param.clone()],
                                            body: body.clone(),
                                            handles: None,
                                            is_async: false,
                                            is_extension: false,
                                            is_overridable: false,
                                            is_overrides: false,
                                            is_must_override: false,
                                            is_shared: false,
                                            is_not_overridable: false,
                                        };
                                        return match self.call_user_sub(&sub, &[val], Some(obj_ref.clone())) {
                                            Ok(_) => Ok(()),
                                            Err(RuntimeError::Exit(ExitType::Sub)) => Ok(()),
                                            Err(e) => Err(e),
                                        };
                                    }
                                }
                            }
                        }
                        
                        // Fallback: Set field directly
                        let member_lower = prop_name.to_lowercase();
                        if val == Value::Nothing {
                        }

                        // Check if this is a BindingSource
                        let obj_type = obj_ref.borrow().fields.get("__type")
                            .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                            .unwrap_or_default();

                        // Handle BindingSource property assignment
                        if obj_type == "BindingSource" {
                            if member_lower == "datasource" {
                                obj_ref.borrow_mut().fields.insert("__datasource".to_string(), val.clone());
                                obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Integer(0));

                                // Propagate DataSourceChanged to all controls bound to this BindingSource
                                let bound: Vec<String> = obj_ref.borrow()
                                    .fields.get("__bound_controls")
                                    .and_then(|v| if let Value::Array(arr) = v {
                                        Some(arr.iter().filter_map(|v| if let Value::String(s) = v { Some(s.clone()) } else { None }).collect())
                                    } else { None })
                                    .unwrap_or_default();
                                let bs_val = Value::Object(obj_ref.clone());
                                for ctrl_name in bound {
                                    let (columns, rows) = self.get_datasource_table_data(&bs_val);
                                    self.side_effects.push_back(crate::RuntimeSideEffect::DataSourceChanged {
                                        control_name: ctrl_name,
                                        columns,
                                        rows,
                                    });
                                }

                                return Ok(());
                            }
                            // Store other BindingSource properties and trigger appropriate side effects
                            if member_lower == "datamember" || member_lower == "filter" || member_lower == "sort" || member_lower == "position" {
                                obj_ref.borrow_mut().fields.insert(member_lower.clone(), val.clone());

                                if member_lower == "position" {
                                    // Position changed → refresh all bound controls
                                    let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                                    let new_pos = val.as_integer().unwrap_or(0);
                                    let count = self.binding_source_row_count_filtered(&Value::Object(obj_ref.clone()));
                                    let bs_name = obj_ref.borrow().fields.get("name")
                                        .map(|v| v.as_string()).unwrap_or_default();
                                    self.side_effects.push_back(crate::RuntimeSideEffect::BindingPositionChanged {
                                        binding_source_name: bs_name,
                                        position: new_pos,
                                        count,
                                    });
                                    self.refresh_bindings_filtered(&obj_ref, &ds, new_pos);
                                } else if member_lower == "filter" || member_lower == "sort" {
                                    // Filter/Sort changed → re-emit DataSourceChanged for all bound controls + refresh bindings
                                    let bound: Vec<String> = obj_ref.borrow()
                                        .fields.get("__bound_controls")
                                        .and_then(|v| if let Value::Array(arr) = v {
                                            Some(arr.iter().filter_map(|v| if let Value::String(s) = v { Some(s.clone()) } else { None }).collect())
                                        } else { None })
                                        .unwrap_or_default();
                                    let bs_val = Value::Object(obj_ref.clone());
                                    for ctrl_name in bound {
                                        let (columns, rows) = self.get_datasource_table_data_filtered(&bs_val);
                                        self.side_effects.push_back(crate::RuntimeSideEffect::DataSourceChanged {
                                            control_name: ctrl_name,
                                            columns,
                                            rows,
                                        });
                                    }
                                    // Reset position to 0 and refresh bindings
                                    obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Integer(0));
                                    let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                                    let count = self.binding_source_row_count_filtered(&Value::Object(obj_ref.clone()));
                                    let bs_name = obj_ref.borrow().fields.get("name")
                                        .map(|v| v.as_string()).unwrap_or_default();
                                    self.side_effects.push_back(crate::RuntimeSideEffect::BindingPositionChanged {
                                        binding_source_name: bs_name,
                                        position: 0,
                                        count,
                                    });
                                    self.refresh_bindings_filtered(&obj_ref, &ds, 0);
                                }
                                return Ok(());
                            }
                        }

                        // Handle DataSource assignment on Object-based controls (DataGridView, etc.)
                        if member_lower == "datasource" {
                            obj_ref.borrow_mut().fields.insert("__datasource".to_string(), val.clone());
                            let mut obj_name: Option<String> = None;
                            if let Some(Value::String(name_val)) = obj_ref.borrow().fields.get("name") {
                                obj_name = Some(name_val.clone());
                            }
                            // Fallback: infer control name from the assignment target expression
                            // when Me.ctrl.Name hasn't been set yet (designer sets DataSource before Name)
                            // e.g. `Me.dgv1.DataSource = Me.bs1` → infer "dgv1" from MemberAccess(Me, "dgv1")
                            if obj_name.is_none() {
                                if let Expression::MemberAccess(inner, ctrl_member) = object {
                                    if matches!(inner.as_ref(), Expression::Me) {
                                        let inferred = ctrl_member.as_str().to_string();
                                        // Store on the object so later assignments can find the name
                                        obj_ref.borrow_mut().fields.insert("name".to_string(), Value::String(inferred.clone()));
                                        obj_name = Some(inferred);
                                    }
                                }
                            }
                            // If the value is a BindingSource, register this control as a subscriber
                            if let Value::Object(bs_ref) = &val {
                                let is_bs = bs_ref.borrow().fields.get("__type")
                                    .and_then(|v| if let Value::String(s) = v { Some(s == "BindingSource") } else { None })
                                    .unwrap_or(false);
                                if is_bs {
                                    if let Some(ref oname) = obj_name {
                                        let mut bs = bs_ref.borrow_mut();
                                        let arr = bs.fields.entry("__bound_controls".to_string())
                                            .or_insert_with(|| Value::Array(Vec::new()));
                                        if let Value::Array(list) = arr {
                                            let name_val = Value::String(oname.clone());
                                            if !list.contains(&name_val) {
                                                list.push(name_val);
                                            }
                                        }
                                    }
                                }
                            }
                            if let Some(oname) = &obj_name {
                                let (columns, rows) = self.get_datasource_table_data(&val);
                                self.side_effects.push_back(crate::RuntimeSideEffect::DataSourceChanged {
                                    control_name: oname.clone(),
                                    columns,
                                    rows,
                                });
                            }
                            return Ok(());
                        }

                        // When assigning a WinForms string-proxy (e.g. New TextBox() → "TextBox")
                        // to an instance field (Me.txt1 = ...), use the field name as the proxy
                        // so that PropertyChange / DataBindings refer to the control by its
                        // instance name ("txt1") rather than the class name ("TextBox").
                        let store_val = if let Value::String(ref s) = val {
                            let s_lower = s.to_lowercase();
                            let is_winforms = matches!(s_lower.as_str(),
                                "textbox" | "label" | "button" | "checkbox" | "radiobutton" |
                                "groupbox" | "panel" | "combobox" | "listbox" | "picturebox" |
                                "timer" | "toolstrip" | "menustrip" | "statusstrip" | "tabcontrol" |
                                "richtextbox" | "progressbar" | "trackbar" | "numericupdown" |
                                "datetimepicker" | "monthcalendar" | "treeview" | "listview" |
                                "webbrowser" | "errorprovider" | "tooltip" |
                                "datagridview" | "bindingnavigator" | "bindingsource" |
                                "flowlayoutpanel" | "tablelayoutpanel" | "splitcontainer" |
                                "maskedtextbox" | "domainupdown" | "contextmenustrip" |
                                "toolstripstatuslabel" | "linklabel" | "hscrollbar" | "vscrollbar" |
                                "checkedlistbox" | "propertygrid" | "splitter" | "datagrid" |
                                "usercontrol" | "toolstripseparator" | "toolstripbutton" |
                                "toolstriplabel" | "toolstripcombobox" | "toolstripdropdownbutton" |
                                "toolstripsplitbutton" | "toolstriptextbox" | "toolstripprogressbar" |
                                "printpreviewcontrol" | "printdialog" | "printpreviewdialog" |
                                "pagesetupdialog" | "helpprovider" | "dataview" |
                                "openfiledialog" | "savefiledialog" | "folderbrowserdialog" |
                                "colordialog" | "fontdialog" | "notifyicon" | "imagelist" |
                                "backgroundworker" | "sqlconnection" | "oledbconnection"
                            );
                            if is_winforms {
                                Value::String(prop_name.clone())
                            } else {
                                val.clone()
                            }
                        } else {
                            val.clone()
                        };
                        obj_ref.borrow_mut().fields.insert(member_lower.clone(), store_val.clone());

                        // StringBuilder.Length setter: truncate or pad the buffer
                        if obj_type == "StringBuilder" && member_lower == "length" {
                            if let Ok(new_len) = store_val.as_integer() {
                                let new_len = new_len as usize;
                                let mut borrow = obj_ref.borrow_mut();
                                if let Some(Value::String(buf)) = borrow.fields.get("__data").cloned() {
                                    let mut chars: Vec<char> = buf.chars().collect();
                                    if new_len < chars.len() {
                                        chars.truncate(new_len);
                                    } else {
                                        chars.resize(new_len, '\0');
                                    }
                                    let new_buf: String = chars.into_iter().collect();
                                    borrow.fields.insert("__data".to_string(), Value::String(new_buf));
                                }
                            }
                        }
                        let is_control = obj_ref.borrow().fields.get("__is_control")
                            .map(|v| v.as_bool().unwrap_or(false)).unwrap_or(false);
                        if is_control {
                            match member_lower.as_str() {
                                "location" => {
                                    // Extract X/Y from Point object and set Left/Top
                                    if let Value::Object(pt) = &val {
                                        let ptb = pt.borrow();
                                        if let Some(xv) = ptb.fields.get("x") {
                                            obj_ref.borrow_mut().fields.insert("left".to_string(), xv.clone());
                                        }
                                        if let Some(yv) = ptb.fields.get("y") {
                                            obj_ref.borrow_mut().fields.insert("top".to_string(), yv.clone());
                                        }
                                    }
                                }
                                "size" => {
                                    // Extract Width/Height from Size object
                                    if let Value::Object(sz) = &val {
                                        let szb = sz.borrow();
                                        if let Some(wv) = szb.fields.get("width") {
                                            obj_ref.borrow_mut().fields.insert("width".to_string(), wv.clone());
                                        }
                                        if let Some(hv) = szb.fields.get("height") {
                                            obj_ref.borrow_mut().fields.insert("height".to_string(), hv.clone());
                                        }
                                    }
                                }
                                "clientsize" => {
                                    if let Value::Object(sz) = &val {
                                        let szb = sz.borrow();
                                        if let Some(wv) = szb.fields.get("width") {
                                            obj_ref.borrow_mut().fields.insert("width".to_string(), wv.clone());
                                        }
                                        if let Some(hv) = szb.fields.get("height") {
                                            obj_ref.borrow_mut().fields.insert("height".to_string(), hv.clone());
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }

                        // If this looks like a control object, push a side-effect so the UI can sync
                        let mut obj_name: Option<String> = None;
                        if let Some(Value::String(name_val)) = obj_ref.borrow().fields.get("name") {
                            obj_name = Some(name_val.clone());
                        }
                        if let Some(oname) = obj_name {
                            self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                                object: oname,
                                property: prop_name,
                                value: val,
                            });
                        }
                        Ok(())
                    }
                    Value::String(obj_name) => {
                        // String proxy (WinForms control)
                        let prop_lower = prop_name.to_lowercase();

                        // DataSource assignment: emit DataSourceChanged for the UI
                        if prop_lower == "datasource" {
                            let ds_key = format!("{}.__datasource", obj_name);
                            self.env.define_global(&ds_key, val.clone());
                            // Resolve table data and push side-effect
                            let (columns, rows) = self.get_datasource_table_data(&val);
                            self.side_effects.push_back(crate::RuntimeSideEffect::DataSourceChanged {
                                control_name: obj_name,
                                columns,
                                rows,
                            });
                            return Ok(());
                        }

                        // BindingSource property setters
                        if prop_lower == "datamember" || prop_lower == "position" || prop_lower == "filter" || prop_lower == "sort" {
                            let key = format!("{}.{}", obj_name, prop_name);
                            self.env.define_global(&key, val.clone());
                            return Ok(());
                        }

                        let key = format!("{}.{}", obj_name, prop_name);
                        self.env.define_global(&key, val.clone());
                        self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                            object: obj_name,
                            property: prop_name,
                            value: val,
                        });
                        Ok(())
                    }
                    _ => {
                        // Fallback for unknown objects - use both identifier and evaluated name
                        // This handles cases like Me.Caption or control references
                        let obj_ident = self.expr_to_string(object);
                        let key = format!("{}.{}", obj_ident, prop_name);
                        self.env.set(&key, val.clone())?;
                        
                        self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                            object: obj_ident,
                            property: prop_name,
                            value: val,
                        });
                        Ok(())
                    }
                }
            }

            Statement::ArrayAssignment { array, indices, value } => {
                let val = self.evaluate_expr(value)?;
                let array_lower = array.as_str().to_lowercase();

                if indices.len() == 1 {
                    // 1-D array assignment
                    let index = self.evaluate_expr(&indices[0])?.as_integer()? as usize;
                    // Try current object fields first
                    if let Some(obj_rc) = &self.current_object {
                        let mut obj = obj_rc.borrow_mut();
                        if let Some(arr_val) = obj.fields.get_mut(&array_lower) {
                            arr_val.set_array_element(index, val.clone())?;
                            return Ok(());
                        }
                    }
                    let mut arr = self.env.get(array.as_str())?;
                    arr.set_array_element(index, val)?;
                    self.env.set(array.as_str(), arr)?;
                } else {
                    // Multi-dimensional assignment: drill into nested arrays
                    let idx_vals: Vec<usize> = indices.iter()
                        .map(|e| self.evaluate_expr(e).and_then(|v| v.as_integer()).map(|i| i as usize))
                        .collect::<Result<Vec<_>, _>>()?;
                    let mut arr = self.env.get(array.as_str())?;
                    set_multi_dim_element(&mut arr, &idx_vals, val)?;
                    self.env.set(array.as_str(), arr)?;
                }
                Ok(())
            }

            Statement::ReDim { preserve, array, bounds } => {
                let new_size = (self.evaluate_expr(&bounds[0])?.as_integer()? + 1) as usize;

                if *preserve {
                    // Get existing array and resize preserving data
                    let mut arr = self.env.get(array.as_str())?;
                    if let Value::Array(ref mut vec) = arr {
                        // Infer element type default from existing elements
                        let default_val = vec.iter().find(|v| !matches!(v, Value::Nothing))
                            .map(|v| match v {
                                Value::Integer(_) => Value::Integer(0),
                                Value::Long(_) => Value::Long(0),
                                Value::Single(_) => Value::Single(0.0),
                                Value::Double(_) => Value::Double(0.0),
                                Value::String(_) => Value::String(String::new()),
                                Value::Boolean(_) => Value::Boolean(false),
                                _ => Value::Nothing,
                            })
                            .unwrap_or(Value::Integer(0));
                        vec.resize(new_size, default_val);
                        self.env.set(array.as_str(), arr)?;
                    } else {
                        return Err(RuntimeError::Custom(format!("{} is not an array", array.as_str())));
                    }
                } else {
                    // Without Preserve: try to infer type from existing array, default to Integer(0)
                    let default_val = self.env.get(array.as_str()).ok()
                        .and_then(|v| if let Value::Array(vec) = v {
                            vec.iter().find(|el| !matches!(el, Value::Nothing)).cloned()
                        } else { None })
                        .map(|v| match v {
                            Value::Integer(_) => Value::Integer(0),
                            Value::Long(_) => Value::Long(0),
                            Value::Single(_) => Value::Single(0.0),
                            Value::Double(_) => Value::Double(0.0),
                            Value::String(_) => Value::String(String::new()),
                            Value::Boolean(_) => Value::Boolean(false),
                            _ => Value::Integer(0),
                        })
                        .unwrap_or(Value::Integer(0));
                    let new_arr = Value::Array(vec![default_val; new_size]);
                    self.env.set(array.as_str(), new_arr)?;
                }
                Ok(())
            }

            Statement::Select { test_expr, cases, else_block } => {
                let test_val = self.evaluate_expr(test_expr)?;
                let mut matched = false;

                for case in cases {
                    // Check if any condition matches
                    for condition in &case.conditions {
                        let matches = match condition {
                            CaseCondition::Value(expr) => {
                                let case_val = self.evaluate_expr(expr)?;
                                values_equal(&test_val, &case_val)
                            }
                            CaseCondition::Range { from, to } => {
                                let from_val = self.evaluate_expr(from)?;
                                let to_val = self.evaluate_expr(to)?;
                                value_in_range(&test_val, &from_val, &to_val)
                            }
                            CaseCondition::Comparison { op, expr } => {
                                let comp_val = self.evaluate_expr(expr)?;
                                compare_values(&test_val, op, &comp_val)?
                            }
                        };

                        if matches {
                            matched = true;
                            break;
                        }
                    }

                    if matched {
                        // Execute case body
                        match self.execute_block(&case.body) {
                            Ok(_) => return Ok(()),
                            Err(RuntimeError::Exit(ExitType::Select)) => return Ok(()),
                            Err(e) => return Err(e),
                        }
                    }
                }

                // If no case matched, execute else block
                if !matched {
                    if let Some(else_stmts) = else_block {
                        self.execute_block(else_stmts)?;
                    }
                }

                Ok(())
            }

            Statement::If {
                condition,
                then_branch,
                elseif_branches,
                else_branch,
            } => {
                let cond_val = self.evaluate_expr(condition)?;

                if cond_val.is_truthy() {
                    for stmt in then_branch {
                        self.execute(stmt)?;
                    }
                } else {
                    let mut executed = false;
                    for (elseif_cond, elseif_stmts) in elseif_branches {
                        let elseif_val = self.evaluate_expr(elseif_cond)?;
                        if elseif_val.is_truthy() {
                            for stmt in elseif_stmts {
                                self.execute(stmt)?;
                            }
                            executed = true;
                            break;
                        }
                    }

                    if !executed {
                        if let Some(else_stmts) = else_branch {
                            for stmt in else_stmts {
                                self.execute(stmt)?;
                            }
                        }
                    }
                }

                Ok(())
            }

            Statement::For {
                variable,
                start,
                end,
                step,
                body,
            } => {
                let start_val = self.evaluate_expr(start)?.as_integer()?;
                let end_val = self.evaluate_expr(end)?.as_integer()?;
                let step_val = if let Some(s) = step {
                    self.evaluate_expr(s)?.as_integer()?
                } else {
                    1
                };

                let mut i = start_val;

                loop {
                    // Check loop condition
                    if (step_val > 0 && i > end_val) || (step_val < 0 && i < end_val) {
                        break;
                    }

                    self.env.set(variable.as_str(), Value::Integer(i))?;

                    // Execute body
                    for stmt in body {
                        match self.execute(stmt) {
                            Err(RuntimeError::Exit(ExitType::For)) => return Ok(()),
                            Err(RuntimeError::Continue(vybe_parser::ast::stmt::ContinueType::For)) => break,
                            Err(e) => return Err(e),
                            Ok(_) => {}
                        }
                    }

                    i += step_val;
                }

                Ok(())
            }

            Statement::While { condition, body } => {
                loop {
                    let cond_val = self.evaluate_expr(condition)?;
                    if !cond_val.is_truthy() {
                        break;
                    }

                    for stmt in body {
                        match self.execute(stmt) {
                            Err(RuntimeError::Exit(ExitType::Do) | RuntimeError::Exit(ExitType::While)) => return Ok(()),
                            // Parser Check: Exit While -> Statement::ExitDo (if mapped) or ExitWhile?
                            // Parser map: "while" -> ???
                            // In parser.rs: "do" -> ExitDo, "for" -> ExitFor.
                            // "while" is not handled in parser.rs exit_statement!
                            // I should add "while" -> ExitWhile support in parser.rs too if VB.NET supports it.
                            // VB.NET supports `Exit While`.
                            // I should check parser.rs exit_statement logic.
                            
                            Err(RuntimeError::Continue(vybe_parser::ast::stmt::ContinueType::While)) => break,
                            Err(e) => return Err(e),
                            Ok(_) => {}
                        }
                    }
                }

                Ok(())
            }

            Statement::DoLoop {
                pre_condition,
                body,
                post_condition,
            } => {
                use vybe_parser::LoopConditionType;

                loop {
                    // Check pre-condition
                    if let Some((cond_type, expr)) = pre_condition {
                        let val = self.evaluate_expr(expr)?;
                        match cond_type {
                            LoopConditionType::While => {
                                if !val.is_truthy() {
                                    break;
                                }
                            }
                            LoopConditionType::Until => {
                                if val.is_truthy() {
                                    break;
                                }
                            }
                        }
                    }

                    // Execute body
                    for stmt in body {
                        match self.execute(stmt) {
                            Err(RuntimeError::Exit(ExitType::Do)) => return Ok(()),
                            Err(RuntimeError::Continue(vybe_parser::ast::stmt::ContinueType::Do)) => break,
                            Err(e) => return Err(e),
                            Ok(_) => {}
                        }
                    }

                    // Check post-condition
                    if let Some((cond_type, expr)) = post_condition {
                        let val = self.evaluate_expr(expr)?;
                        match cond_type {
                            LoopConditionType::While => {
                                if !val.is_truthy() {
                                    break;
                                }
                            }
                            LoopConditionType::Until => {
                                if val.is_truthy() {
                                    break;
                                }
                            }
                        }
                    } else if pre_condition.is_none() {
                        // Infinite loop without condition
                        // In practice, there should be an Exit Do
                    }
                }

                Ok(())
            }

            Statement::ExitSub => Err(RuntimeError::Exit(ExitType::Sub)),
            Statement::ExitFunction => Err(RuntimeError::Exit(ExitType::Function)),
            Statement::ExitFor => Err(RuntimeError::Exit(ExitType::For)),
            Statement::ExitDo => Err(RuntimeError::Exit(ExitType::Do)),
            Statement::ExitWhile => Err(RuntimeError::Exit(ExitType::While)),
            Statement::ExitSelect => Err(RuntimeError::Exit(ExitType::Select)),
            Statement::ExitTry => return Err(RuntimeError::Exit(ExitType::Try)),
            
            Statement::Continue(typ) => return Err(RuntimeError::Continue(typ.clone())),

            Statement::Throw(expr) => {
                if let Some(ex) = expr {
                    let val = self.evaluate_expr(ex)?;
                    // If the thrown value is an exception object, extract type + message
                    if let Value::Object(ref obj) = val {
                        let b = obj.borrow();
                        let ex_type = b.class_name.clone();
                        let msg = b.fields.get("message")
                            .map(|v| v.as_string())
                            .unwrap_or_else(|| format!("{}", ex_type));
                        let inner = b.fields.get("innerexception")
                            .and_then(|v| if *v == Value::Nothing { None } else { Some(v.as_string()) });
                        return Err(RuntimeError::Exception(ex_type, msg, inner));
                    }
                    // Plain string throw
                    return Err(RuntimeError::Exception("Exception".to_string(), val.as_string(), None));
                } else {
                    // Re-throw current exception (Throw without expression)
                    return Err(RuntimeError::Exception("Exception".to_string(), "An exception was thrown".to_string(), None));
                }
            }

            Statement::AddHandler { event_target, handler } => {
                // event_target = "Button1.Click", handler = "HandleClick" or "Me.HandleClick"
                let parts: Vec<&str> = event_target.splitn(2, '.').collect();
                if parts.len() == 2 {
                    let control = parts[0].to_string();
                    let event = parts[1].to_string();
                    // Resolve the handler name: strip "Me." prefix if present
                    let handler_name = if handler.to_lowercase().starts_with("me.") {
                        handler[3..].to_string()
                    } else {
                        handler.clone()
                    };

                    // Check if the control is a BackgroundWorker object — store handler on object
                    let event_lower = event.to_lowercase();
                    let is_bgw = if let Ok(val) = self.env.get(&control) {
                        if let Value::Object(ref obj_ref) = val {
                            let tn = obj_ref.borrow().fields.get("__type").map(|v| v.as_string()).unwrap_or_default();
                            if tn == "BackgroundWorker" {
                                let field_key = match event_lower.as_str() {
                                    "dowork" => Some("__dowork_handler"),
                                    "progresschanged" => Some("__progresschanged_handler"),
                                    "runworkercompleted" => Some("__runworkercompleted_handler"),
                                    _ => None,
                                };
                                if let Some(key) = field_key {
                                    obj_ref.borrow_mut().fields.insert(key.to_string(), Value::String(handler_name.clone()));
                                }
                                true
                            } else { false }
                        } else { false }
                    } else { false };

                    if !is_bgw {
                        // Register in the event system (WinForms)
                        if let Some(event_type) = vybe_forms::EventType::from_name(&event) {
                            self.events.register_handler(&control, &event_type, &handler_name);
                        }
                    }
                }
                Ok(())
            }

            Statement::RemoveHandler { event_target, handler } => {
                let parts: Vec<&str> = event_target.splitn(2, '.').collect();
                if parts.len() == 2 {
                    let control = parts[0].to_string();
                    let event = parts[1].to_string();
                    let handler_name = if handler.to_lowercase().starts_with("me.") {
                        handler[3..].to_string()
                    } else {
                        handler.clone()
                    };
                    if let Some(event_type) = vybe_forms::EventType::from_name(&event) {
                        self.events.remove_handler(&control, &event_type, &handler_name);
                    }
                }
                Ok(())
            }

            Statement::Try { body, catches, finally } => {
                let result = self.execute_block(body);
                
                let mut flow_result = result;
                
                if let Err(e) = &flow_result {
                    // Check if it's a runtime error (not Exit/Continue/Return unless handled?)
                    // VB.NET catches Exceptions. Exit/Return are control flow, usually not caught by Catch (unless catch all?).
                    // But implemented as Result::Err in Rust.
                    
                    let mut handled = false;
                    
                    // Only catch actual value errors, not control flow (Exit, Return, Continue)
                    let is_error = match e {
                        RuntimeError::Exit(_) | RuntimeError::Return(_) | RuntimeError::Continue(_) => false,
                        _ => true,
                    };

                    if is_error {
                        // Extract exception type from the error
                        let (ex_type, ex_msg, ex_inner) = match e {
                            RuntimeError::Exception(t, m, inner) => (t.clone(), m.clone(), inner.clone()),
                            RuntimeError::Custom(m) => ("Exception".to_string(), m.clone(), None),
                            RuntimeError::TypeError { expected, got } => ("TypeMismatchException".to_string(), format!("Type error: expected {}, got {}", expected, got), None),
                            RuntimeError::UndefinedVariable(v) => ("NullReferenceException".to_string(), format!("Undefined variable: {}", v), None),
                            RuntimeError::UndefinedFunction(f) => ("MissingMethodException".to_string(), format!("Undefined function: {}", f), None),
                            RuntimeError::DivisionByZero => ("DivideByZeroException".to_string(), "Division by zero".to_string(), None),
                            _ => ("Exception".to_string(), format!("{}", e), None),
                        };

                        for catch in catches {
                             // Check variable type match
                             let type_match = if let Some((_, Some(type_name))) = &catch.variable {
                                 let catch_type = format!("{:?}", type_name).to_lowercase();
                                 let catch_type = catch_type.trim_start_matches("custom(\"").trim_end_matches("\")");
                                 let ex_lower = ex_type.to_lowercase();
                                 catch_type == "exception" || catch_type == "system.exception"
                                     || ex_lower == catch_type
                                     || ex_lower.ends_with(&format!(".{}", catch_type))
                                     || catch_type.ends_with(&ex_lower)
                                     || ex_lower.ends_with("exception")
                             } else {
                                 true // Catch All
                             };
                             
                             if type_match {
                                 // Push scope for the catch variable to exist during 'When' evaluation and body execution
                                 self.env.push_scope();
                                 
                                 // Define the exception variable if present
                                 if let Some((name, _)) = &catch.variable {
                                     let mut ex_fields = std::collections::HashMap::new();
                                     ex_fields.insert("message".to_string(), Value::String(ex_msg.clone()));
                                     // Basic stacktrace support
                                     ex_fields.insert("stacktrace".to_string(), Value::String(String::new()));
                                     ex_fields.insert("source".to_string(), Value::String(String::new()));
                                     ex_fields.insert("hresult".to_string(), Value::Integer(-2146233088));
                                     if let Some(ref inner) = ex_inner {
                                         ex_fields.insert("innerexception".to_string(), Value::String(inner.clone()));
                                     } else {
                                          ex_fields.insert("innerexception".to_string(), Value::Nothing);
                                     }
                                     
                                     let ex_obj = crate::value::ObjectData {
                                         class_name: ex_type.clone(),
                                         fields: ex_fields,
                                     };
                                     self.env.define(name.as_str(), Value::Object(std::rc::Rc::new(std::cell::RefCell::new(ex_obj))));
                                 }

                                 let when_match = if let Some(expr) = &catch.when_clause {
                                      match self.evaluate_expr(expr) {
                                          Ok(val) => val.is_truthy(),
                                          Err(_) => false, // If evaluation fails (e.g. wrong property), treat as false match
                                      }
                                 } else {
                                      true
                                 };
                                 
                                 if when_match {
                                     flow_result = self.execute_block(&catch.body);
                                     self.env.pop_scope();
                                     handled = true;
                                     break;
                                 } else {
                                     // Pop scope if we didn't match (remove the variable)
                                     self.env.pop_scope();
                                 }
                             }
                        }
                        
                        let _ = handled;
                    }
                }
                
                // Finally block
                if let Some(final_stmts) = finally {
                     let final_res = self.execute_block(final_stmts);
                     if final_res.is_err() {
                         flow_result = final_res;
                     }
                }
                
                return flow_result;
            }

            Statement::Return(value) => {
                let val = if let Some(expr) = value {
                    Some(self.evaluate_expr(expr)?)
                } else {
                    None
                };
                Err(RuntimeError::Return(val))
            }

            Statement::Call { name, arguments } => {
                // If the call name contains dots (e.g. "Me.txt1.DataBindings.Add"),
                // the parser failed to produce an expression chain. Build and evaluate
                // a synthetic MethodCall expression to handle it correctly.
                if name.as_str().contains('.') {
                    let parts: Vec<&str> = name.as_str().split('.').collect();
                    if parts.len() >= 2 {
                        let mut expr: Expression = if parts[0].eq_ignore_ascii_case("me") {
                            Expression::Me
                        } else {
                            Expression::Variable(Identifier::new(parts[0]))
                        };
                        for part in &parts[1..parts.len()-1] {
                            expr = Expression::MemberAccess(Box::new(expr), Identifier::new(*part));
                        }
                        let method_name = Identifier::new(parts[parts.len()-1]);
                        let call_expr = Expression::MethodCall(
                            Box::new(expr),
                            method_name,
                            arguments.to_vec(),
                        );
                        self.evaluate_expr(&call_expr)?;
                        return Ok(());
                    }
                }
                self.call_procedure(name, arguments)?;
                Ok(())
            }

            Statement::ExpressionStatement(expr) => {
                self.evaluate_expr(expr)?;
                Ok(())
            }

            Statement::ForEach { variable, collection, body } => {
                let coll_val = self.evaluate_expr(collection)?;
                let items = coll_val.to_iterable()?;
                self.env.define(variable.as_str(), Value::Nothing);
                for item in &items {
                    self.env.set(variable.as_str(), item.clone())?;
                    let mut should_exit = false;
                    for s in body {
                        match self.execute(s) {
                            Err(RuntimeError::Exit(ExitType::For)) => { should_exit = true; break; }
                            Err(RuntimeError::Continue(vybe_parser::ast::stmt::ContinueType::For)) => break,
                            Err(e) => return Err(e),
                            Ok(()) => {}
                        }
                    }
                    if should_exit { break; }
                }
                Ok(())
            }

            Statement::With { object, body } => {
                let obj_val = self.evaluate_expr(object)?;
                let prev_with = self.with_object.take();
                self.with_object = Some(obj_val);
                for s in body {
                    self.execute(s)?;
                }
                self.with_object = prev_with;
                Ok(())
            }

            Statement::Using { variable, resource, body } => {
                // Evaluate the resource and bind to the given variable for the scope of the block.
                let res_val = self.evaluate_expr(resource)?;
                self.env.define(variable.as_str(), res_val.clone());

                for s in body {
                    self.execute(s)?;
                }

                // Best-effort dispose: if the object has a Dispose method, invoke it.
                if let Value::Object(obj_ref) = res_val.clone() {
                    let class_name = obj_ref.borrow().class_name.clone();
                    if let Some(method) = self.find_method(&class_name, "Dispose") {
                        if let vybe_parser::ast::decl::MethodDecl::Sub(s) = method {
                            let _ = self.call_user_sub(&s, &[], Some(obj_ref.clone()));
                        }
                    }
                }

                // Clear the variable after leaving the Using scope.
                let _ = self.env.set(variable.as_str(), Value::Nothing);
                Ok(())
            }

            Statement::Open { file_path, mode, file_number } => {
                let path = self.evaluate_expr(file_path)?.as_string();
                let fn_val = self.evaluate_expr(file_number)?.as_integer()?;
                let handle = crate::file_io::open_file(&path, *mode)?;
                self.file_handles.insert(fn_val, handle);
                Ok(())
            }

            Statement::CloseFile { file_number } => {
                if let Some(fn_expr) = file_number {
                    let fn_val = self.evaluate_expr(fn_expr)?.as_integer()?;
                    self.file_handles.remove(&fn_val);
                } else {
                    self.file_handles.clear();
                }
                Ok(())
            }

            Statement::PrintFile { file_number, items, newline } => {
                let fn_val = self.evaluate_expr(file_number)?.as_integer()?;
                
                // Evaluate all items first to avoid double borrow
                let mut values_to_print = Vec::new();
                for item in items {
                    values_to_print.push(self.evaluate_expr(item)?.as_string());
                }

                if let Some(handle) = self.file_handles.get_mut(&fn_val) {
                     for (i, s) in values_to_print.iter().enumerate() {
                         crate::file_io::print_string(handle, s)?;
                         if i < values_to_print.len() - 1 {
                             crate::file_io::print_string(handle, "\t")?;
                         }
                     }
                     if *newline {
                         crate::file_io::write_line(handle, "")?;
                     }
                } else {
                    return Err(RuntimeError::Custom(format!("File number {} not open", fn_val)));
                }
                Ok(())
            }

            Statement::WriteFile { file_number, items } => {
                let fn_val = self.evaluate_expr(file_number)?.as_integer()?;
                
                // Evaluate all items first
                let mut values_to_write = Vec::new();
                for item in items {
                    let val = self.evaluate_expr(item)?;
                    let s = match val {
                        Value::String(s) => format!("\"{}\"", s),
                        _ => val.as_string(),
                    };
                    values_to_write.push(s);
                }

                if let Some(handle) = self.file_handles.get_mut(&fn_val) {
                     crate::file_io::write_line(handle, &values_to_write.join(","))?;
                } else {
                    return Err(RuntimeError::Custom(format!("File number {} not open", fn_val)));
                }
                Ok(())
            }

            Statement::InputFile { file_number, variables } => {
                 let fn_val = self.evaluate_expr(file_number)?.as_integer()?;
                 
                 // Read line first
                 let line = if let Some(handle) = self.file_handles.get_mut(&fn_val) {
                     crate::file_io::read_line(handle)?
                 } else {
                     return Err(RuntimeError::Custom(format!("File number {} not open", fn_val)));
                 };

                 // Parse and set variables
                 let parts: Vec<&str> = line.split(',').collect();
                 for (i, var) in variables.iter().enumerate() {
                     if i < parts.len() {
                         let val_str = parts[i].trim().trim_matches('"');
                         let val = if let Ok(int_val) = val_str.parse::<i32>() {
                             Value::Integer(int_val)
                         } else {
                             Value::String(val_str.to_string())
                         };
                         self.env.set(var.as_str(), val)?;
                     }
                 }
                Ok(())
            }

            Statement::LineInput { file_number, variable } => {
                let fn_val = self.evaluate_expr(file_number)?.as_integer()?;
                
                // Read line first
                let line = if let Some(handle) = self.file_handles.get_mut(&fn_val) {
                     crate::file_io::read_line(handle)?
                } else {
                    return Err(RuntimeError::Custom(format!("File number {} not open", fn_val)));
                };
                
                self.env.set(variable.as_str(), Value::String(line))?;
                Ok(())
            }
            Statement::ExitProperty => {
                 return Err(RuntimeError::Exit(ExitType::Property));
            }
            Statement::CompoundAssignment { target, members, indices, operator, value } => {
                // Get current value
                let current = if !members.is_empty() {
                    // Build member access expression
                    let mut obj = Expression::Variable(target.clone());
                    for m in members.iter() {
                        obj = Expression::MemberAccess(Box::new(obj), m.clone());
                    }
                    self.evaluate_expr(&obj)?
                } else if !indices.is_empty() {
                    let arr_expr = Expression::ArrayAccess(target.clone(), indices.clone());
                    self.evaluate_expr(&arr_expr)?
                } else {
                    self.evaluate_expr(&Expression::Variable(target.clone()))?
                };

                let rhs = self.evaluate_expr(value)?;
                use vybe_parser::ast::stmt::CompoundOp;
                let new_val = match operator {
                    CompoundOp::AddAssign => {
                        match (&current, &rhs) {
                            (Value::Integer(a), Value::Integer(b)) => Value::Integer(a + b),
                            (Value::String(a), _) => Value::String(format!("{}{}", a, rhs.as_string())),
                            _ => Value::Double(current.as_double()? + rhs.as_double()?),
                        }
                    }
                    CompoundOp::SubtractAssign => {
                        match (&current, &rhs) {
                            (Value::Integer(a), Value::Integer(b)) => Value::Integer(a - b),
                            _ => Value::Double(current.as_double()? - rhs.as_double()?),
                        }
                    }
                    CompoundOp::MultiplyAssign => {
                        match (&current, &rhs) {
                            (Value::Integer(a), Value::Integer(b)) => Value::Integer(a * b),
                            _ => Value::Double(current.as_double()? * rhs.as_double()?),
                        }
                    }
                    CompoundOp::DivideAssign => {
                        let a = current.as_double()?;
                        let b = rhs.as_double()?;
                        if b == 0.0 { return Err(RuntimeError::DivisionByZero); }
                        Value::Double(a / b)
                    }
                    CompoundOp::IntDivideAssign => {
                        let a = current.as_integer()?;
                        let b = rhs.as_integer()?;
                        if b == 0 { return Err(RuntimeError::DivisionByZero); }
                        Value::Integer(a / b)
                    }
                    CompoundOp::ConcatAssign => {
                        Value::String(format!("{}{}", current.as_string(), rhs.as_string()))
                    }
                    CompoundOp::ExponentAssign => {
                        Value::Double(current.as_double()?.powf(rhs.as_double()?))
                    }
                    CompoundOp::ShiftLeftAssign => {
                        Value::Long(current.as_long()? << rhs.as_integer()? as u32)
                    }
                    CompoundOp::ShiftRightAssign => {
                        Value::Long(current.as_long()? >> rhs.as_integer()? as u32)
                    }
                };

                // Assign back
                if !members.is_empty() {
                    // For member assignment, build the path
                    let full_name = std::iter::once(target.as_str().to_string())
                        .chain(members.iter().map(|m| m.as_str().to_string()))
                        .collect::<Vec<_>>()
                        .join(".");
                    self.env.set(&full_name, new_val)?;
                } else if !indices.is_empty() {
                    let idx = self.evaluate_expr(&indices[0])?.as_integer()? as usize;
                    if let Ok(Value::Array(mut arr)) = self.env.get(target.as_str()) {
                        if idx < arr.len() {
                            arr[idx] = new_val.clone();
                            self.env.set(target.as_str(), Value::Array(arr))?;
                        }
                    }
                } else {
                    self.env.set(target.as_str(), new_val)?;
                }
                Ok(())
            }
            Statement::RaiseEvent { event_name, arguments } => {
                // RaiseEvent fires events registered via AddHandler
                // Look up handler in the event system
                let event_str = event_name.as_str();
                // Try to find a handler: use current module or "Me" as context
                let module = self.current_module.clone().unwrap_or_default();
                // The handler is typically registered as "ControlName_EventName" 
                // For form-level events, the control is the form itself
                if let Some(event_type) = vybe_forms::EventType::from_name(event_str) {
                    // Evaluate args first
                    let args: Vec<Value> = arguments.iter()
                        .map(|a| self.evaluate_expr(a))
                        .collect::<Result<Vec<_>, _>>()?;
                        
                    // Get handlers
                    let handlers = if let Some(h_vec) = self.events.get_handlers(&module, &event_type) {
                        h_vec.clone()
                    } else {
                        Vec::new()
                    };
                    
                    for handler_name in handlers {
                        self.call_event_handler(&handler_name, &args)?;
                    }
                }
                Ok(())
            }
            // --- Static local variables ---
            Statement::StaticVar { name, var_type, initializer } => {
                // Static variables persist across calls. Key by module + procedure + var name.
                let module = self.current_module.clone().unwrap_or_else(|| "__global__".to_string());
                let proc = self.current_procedure.clone().unwrap_or_else(|| "__main__".to_string());
                let key = format!("{}.{}.{}", module, proc, name.as_str()).to_lowercase();
                if !self.static_locals.contains_key(&key) {
                    let val = if let Some(init) = initializer {
                        self.evaluate_expr(init)?
                    } else {
                        default_value_for_type(name.as_str(), var_type)
                    };
                    self.static_locals.insert(key.clone(), val.clone());
                    self.env.define(name.as_str(), val);
                } else {
                    // Already initialized — just define in current scope with persisted value
                    let val = self.static_locals.get(&key).cloned().unwrap_or(Value::Nothing);
                    self.env.define(name.as_str(), val);
                }
                Ok(())
            }
            // --- GoTo / Label / On Error ---
            Statement::GoTo(label) => {
                Err(RuntimeError::GoTo(label.clone()))
            }
            Statement::Label(_label) => {
                // Labels are markers — execution just passes through them.
                Ok(())
            }
            Statement::OnErrorResumeNext => {
                self.on_error_resume_next = true;
                self.on_error_goto_label = None;
                Ok(())
            }
            Statement::OnErrorGoTo(label) => {
                if label == "0" {
                    // On Error GoTo 0 → disable error handling
                    self.on_error_resume_next = false;
                    self.on_error_goto_label = None;
                } else {
                    self.on_error_resume_next = false;
                    self.on_error_goto_label = Some(label.clone());
                }
                Ok(())
            }
            Statement::Resume(target) => {
                // Resume is only meaningful inside an error handler.
                // For now, just return Ok — the GoTo mechanism handles jumps.
                match target {
                    vybe_parser::ast::stmt::ResumeTarget::Next => Ok(()),
                    vybe_parser::ast::stmt::ResumeTarget::Label(lbl) => Err(RuntimeError::GoTo(lbl.clone())),
                    vybe_parser::ast::stmt::ResumeTarget::Implicit => Ok(()),
                }
            }
        }
    }

    /// Execute a block of statements with GoTo jump support and On Error handling.
    /// When a GoTo is encountered, finds the target Label in the body and resumes
    /// execution from there. Also handles On Error Resume Next (swallow errors).
    fn execute_body_with_goto(&mut self, body: &[Statement]) -> Result<(), RuntimeError> {
        let mut pc = 0; // program counter — index into body
        while pc < body.len() {
            let stmt = &body[pc];
            match self.execute(stmt) {
                Ok(_) => {
                    pc += 1;
                }
                Err(RuntimeError::GoTo(label)) => {
                    // Find the label in the body and jump to it
                    if let Some(idx) = body.iter().position(|s| matches!(s, Statement::Label(l) if l.eq_ignore_ascii_case(&label))) {
                        pc = idx + 1; // resume after the label
                    } else {
                        return Err(RuntimeError::Custom(format!("Label '{}' not found", label)));
                    }
                }
                Err(e) => {
                    // On Error Resume Next: swallow the error and continue
                    if self.on_error_resume_next {
                        pc += 1;
                        continue;
                    }
                    // On Error GoTo <label>: jump to error handler label
                    if let Some(ref lbl) = self.on_error_goto_label.clone() {
                        if let Some(idx) = body.iter().position(|s| matches!(s, Statement::Label(l) if l.eq_ignore_ascii_case(lbl))) {
                            // Disable error handling to avoid infinite loops in the handler
                            self.on_error_goto_label = None;
                            pc = idx + 1;
                            continue;
                        }
                    }
                    // Propagate other errors (Exit Sub/Function, Return, etc.)
                    return Err(e);
                }
            }
        }
        Ok(())
    }

    /// Persist any static local variables from the current scope back into
    /// `self.static_locals` so they survive across calls.
    fn persist_static_locals(&mut self) {
        let module = self.current_module.clone().unwrap_or_else(|| "__global__".to_string());
        let proc = self.current_procedure.clone().unwrap_or_else(|| "__main__".to_string());
        let prefix = format!("{}.{}.", module, proc).to_lowercase();
        let keys: Vec<String> = self.static_locals.keys()
            .filter(|k| k.starts_with(&prefix))
            .cloned()
            .collect();
        for key in keys {
            // Extract var name from key
            let var_name = &key[prefix.len()..];
            if let Ok(val) = self.env.get(var_name) {
                self.static_locals.insert(key, val);
            }
        }
    }

    pub fn evaluate_expr(&mut self, expr: &Expression) -> Result<Value, RuntimeError> {
        match expr {
            Expression::Lambda { params, body, .. } => {
                Ok(Value::Lambda {
                    params: params.clone(),
                    body: body.clone(),
                    env: Rc::new(RefCell::new(self.env.clone())),
                })
            }
            Expression::Call(name, args) => {
                // 1. Check if name refers to a variable holding a Lambda or Array
                // NOTE: We check the environment directly instead of using evaluate_expr(Variable(...))
                // because Variable evaluation has an implicit function-call fallback that would
                // execute the sub/function a first time, then call_function below would execute it again.
                let var_name = name.as_str();
                if let Ok(val) = self.env.get(var_name) {
                   match val {
                       Value::Lambda { .. } => {
                           let arg_values: Result<Vec<_>, _> = args.iter().map(|e| self.evaluate_expr(e)).collect();
                           return self.call_lambda(val, &arg_values?);
                       }
                       Value::Array(arr) => {
                           // Array access via Call syntax — supports multi-dimensional
                           if args.len() == 1 {
                               let index = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                               return arr.get(index).cloned().ok_or_else(|| RuntimeError::Custom("Array index out of bounds".to_string()));
                           } else {
                               // Multi-dim: drill into nested arrays
                               let mut current = Value::Array(arr);
                               for idx_expr in args.iter() {
                                   let idx = self.evaluate_expr(idx_expr)?.as_integer()? as usize;
                                   current = match current {
                                       Value::Array(inner) => inner.get(idx).cloned().ok_or_else(|| {
                                           RuntimeError::Custom("Array index out of bounds".to_string())
                                       })?,
                                       _ => return Err(RuntimeError::Custom("Cannot index non-array dimension".to_string())),
                                   };
                               }
                               return Ok(current);
                           }
                       }
                       Value::Dictionary(dict) => {
                           // Dictionary access via Call syntax (e.g. dict("key"))
                           if args.len() != 1 {
                               return Err(RuntimeError::Custom("Dictionary index must be 1 key".to_string()));
                           }
                           let key = self.evaluate_expr(&args[0])?;
                           return dict.borrow().item(&key);
                       }
                       Value::Collection(col) => {
                           // Collection access via Call syntax (e.g. col("key") or col(0))
                           if args.len() == 1 {
                               let idx_val = self.evaluate_expr(&args[0])?;
                               return match &idx_val {
                                   Value::String(key) => col.borrow().item_by_key(key),
                                   _ => {
                                       let idx = idx_val.as_integer()? as usize;
                                       col.borrow().item(idx)
                                   }
                               };
                           }
                       }
                       _ => {} // Not a callable value, proceed to function lookup
                   }
                }
                
                // 2. Standard function lookup
                self.call_function(name, args)
            }
            Expression::Await(operand) => {
                // Await evaluates the operand; if it's a Task, return its Result
                let val = self.evaluate_expr(operand)?;
                if let Value::Object(ref obj) = val {
                    let b = obj.borrow();
                    if b.class_name == "Task" {
                        // Check if faulted
                        if let Some(Value::Boolean(true)) = b.fields.get("isfaulted") {
                            let msg = b.fields.get("exception").map(|v| v.as_string()).unwrap_or_else(|| "Task faulted".to_string());
                            return Err(RuntimeError::Exception("AggregateException".to_string(), msg, None));
                        }
                        return Ok(b.fields.get("result").cloned().unwrap_or(Value::Nothing));
                    }
                }
                Ok(val)
            }
            Expression::MethodCall(obj, method, args) => {
                match self.call_method(obj, method, args) {
                    Ok(val) => Ok(val),
                    Err(e) => {
                        // Fallback: try extension methods
                        if let Ok(result) = self.try_extension_method(obj, method, args) {
                            Ok(result)
                        } else {
                            Err(e)
                        }
                    }
                }
            }
            Expression::ArrayAccess(array, indices) => {
                let arr_val = self.env.get(array.as_str())?;
                match arr_val {
                    Value::Array(arr) => {
                        if indices.len() == 1 {
                            let index = self.evaluate_expr(&indices[0])?.as_integer()? as usize;
                            arr.get(index).cloned().ok_or_else(|| RuntimeError::Custom("Array index out of bounds".to_string()))
                        } else {
                            // Multi-dimensional: drill into nested arrays
                            let mut current = Value::Array(arr);
                            for idx_expr in indices {
                                let idx = self.evaluate_expr(idx_expr)?.as_integer()? as usize;
                                current = match current {
                                    Value::Array(inner) => inner.get(idx).cloned().ok_or_else(|| {
                                        RuntimeError::Custom("Array index out of bounds".to_string())
                                    })?,
                                    _ => return Err(RuntimeError::Custom("Cannot index non-array dimension".to_string())),
                                };
                            }
                            Ok(current)
                        }
                    }
                    Value::Collection(col) => {
                        let idx_val = self.evaluate_expr(&indices[0])?;
                        match &idx_val {
                            Value::String(key) => {
                                col.borrow().item_by_key(key)
                            }
                            _ => {
                                // VB.NET Collection is 1-based, but ArrayList is 0-based.
                                // We use 0-based here since our Collection wraps ArrayList.
                                let index = idx_val.as_integer()? as usize;
                                col.borrow().item(index)
                            }
                        }
                    }
                    Value::Dictionary(dict) => {
                        if indices.len() != 1 {
                            return Err(RuntimeError::Custom("Dictionary index must be 1 key".to_string()));
                        }
                        let key = self.evaluate_expr(&indices[0])?;
                        return dict.borrow().item(&key);
                    }
                    _ => Err(RuntimeError::Custom(format!("Type is not indexable: {:?}", arr_val))),
                }
            }
            Expression::ArrayLiteral(elements) => {
                let vals: Result<Vec<Value>, RuntimeError> = elements
                    .iter()
                    .map(|e| self.evaluate_expr(e))
                    .collect();
                Ok(Value::Array(vals?))
            }
            Expression::Query(query) => self.execute_query(query),
            Expression::XmlLiteral(node) => self.construct_xml(node),
            // For all other expressions that might contain nested calls, evaluate through interpreter
            Expression::Add(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                match (&l, &r) {
                    (Value::Date(d), Value::Double(n)) => Ok(Value::Date(d + n)),
                    (Value::Double(n), Value::Date(d)) => Ok(Value::Date(d + n)),
                    (Value::Date(d), Value::Integer(n)) => Ok(Value::Date(d + *n as f64)),
                    (Value::Integer(n), Value::Date(d)) => Ok(Value::Date(d + *n as f64)),
                    (Value::Integer(a), Value::Integer(b)) => Ok(Value::Integer(a + b)),
                    _ => {
                        let a = l.as_double()?;
                        let b = r.as_double()?;
                        Ok(Value::Double(a + b))
                    }
                }
            }
            Expression::Subtract(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                match (&l, &r) {
                    (Value::Date(d), Value::Double(n)) => Ok(Value::Date(d - n)),
                    (Value::Date(d1), Value::Date(d2)) => Ok(Value::Double(d1 - d2)),
                    (Value::Date(d), Value::Integer(n)) => Ok(Value::Date(d - *n as f64)),
                    (Value::Integer(a), Value::Integer(b)) => Ok(Value::Integer(a - b)),
                    _ => {
                        let a = l.as_double()?;
                        let b = r.as_double()?;
                        Ok(Value::Double(a - b))
                    }
                }
            }
            Expression::Multiply(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                match (&l, &r) {
                    (Value::Integer(a), Value::Integer(b)) => Ok(Value::Integer(a * b)),
                    _ => {
                        let a = l.as_double()?;
                        let b = r.as_double()?;
                        Ok(Value::Double(a * b))
                    }
                }
            }
            Expression::Divide(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let a = l.as_double()?;
                let b = r.as_double()?;
                if b == 0.0 {
                    return Err(RuntimeError::DivisionByZero);
                }
                Ok(Value::Double(a / b))
            }
            Expression::IntegerDivide(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let a = l.as_integer()?;
                let b = r.as_integer()?;
                if b == 0 {
                    return Err(RuntimeError::DivisionByZero);
                }
                Ok(Value::Integer(a / b))
            }
            Expression::Exponent(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                Ok(Value::Double(l.as_double()?.powf(r.as_double()?)))
            }
            Expression::Concatenate(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                Ok(Value::String(format!("{}{}", l.as_string(), r.as_string())))
            }
            Expression::Negate(operand) => {
                let val = self.evaluate_expr(operand)?;
                match val {
                    Value::Integer(i) => Ok(Value::Integer(-i)),
                    Value::Double(d) => Ok(Value::Double(-d)),
                    _ => {
                        let d = val.as_double()?;
                        Ok(Value::Double(-d))
                    }
                }
            }
            Expression::Not(operand) => {
                let val = self.evaluate_expr(operand)?;
                match val {
                    Value::Boolean(b) => Ok(Value::Boolean(!b)),
                    _ => {
                        let i = val.as_long()?;
                        Ok(Value::Long(!i))
                    }
                }
            }
            // Comparison operators - must be handled here so inner expressions 
            // (like obj field accesses) resolve via the interpreter context
            Expression::Equal(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                Ok(Value::Boolean(values_equal(&l, &r)))
            }
            Expression::NotEqual(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                Ok(Value::Boolean(!values_equal(&l, &r)))
            }
            Expression::LessThan(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let result = match (&l, &r) {
                    (Value::Integer(a), Value::Integer(b)) => a < b,
                    (Value::String(a), Value::String(b)) => a < b,
                    (Value::Date(a), Value::Date(b)) => a < b,
                    _ => l.as_double()? < r.as_double()?,
                };
                Ok(Value::Boolean(result))
            }
            Expression::LessThanOrEqual(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let result = match (&l, &r) {
                    (Value::Integer(a), Value::Integer(b)) => a <= b,
                    (Value::String(a), Value::String(b)) => a <= b,
                    (Value::Date(a), Value::Date(b)) => a <= b,
                    _ => l.as_double()? <= r.as_double()?,
                };
                Ok(Value::Boolean(result))
            }
            Expression::GreaterThan(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let result = match (&l, &r) {
                    (Value::Integer(a), Value::Integer(b)) => a > b,
                    (Value::String(a), Value::String(b)) => a > b,
                    (Value::Date(a), Value::Date(b)) => a > b,
                    _ => l.as_double()? > r.as_double()?,
                };
                Ok(Value::Boolean(result))
            }
            Expression::GreaterThanOrEqual(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let result = match (&l, &r) {
                    (Value::Integer(a), Value::Integer(b)) => a >= b,
                    (Value::String(a), Value::String(b)) => a >= b,
                    (Value::Date(a), Value::Date(b)) => a >= b,
                    _ => l.as_double()? >= r.as_double()?,
                };
                Ok(Value::Boolean(result))
            }
            Expression::And(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                match (&l, &r) {
                    (Value::Boolean(a), Value::Boolean(b)) => Ok(Value::Boolean(*a && *b)),
                    _ => {
                        let i_l = l.as_long()?;
                        let i_r = r.as_long()?;
                        Ok(Value::Long(i_l & i_r))
                    }
                }
            }
            Expression::Or(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                match (&l, &r) {
                    (Value::Boolean(a), Value::Boolean(b)) => Ok(Value::Boolean(*a || *b)),
                    _ => {
                        let i_l = l.as_long()?;
                        let i_r = r.as_long()?;
                        Ok(Value::Long(i_l | i_r))
                    }
                }
            }
            Expression::Xor(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                match (&l, &r) {
                    (Value::Boolean(a), Value::Boolean(b)) => Ok(Value::Boolean(*a ^ *b)),
                    _ => {
                        let i_l = l.as_long()?;
                        let i_r = r.as_long()?;
                        Ok(Value::Long(i_l ^ i_r))
                    }
                }
            }
            Expression::AndAlso(left, right) => {
                // Short-circuit: if left is false, don't evaluate right
                let l = self.evaluate_expr(left)?;
                if !l.as_bool()? {
                    return Ok(Value::Boolean(false));
                }
                let r = self.evaluate_expr(right)?;
                Ok(Value::Boolean(r.as_bool()?))
            }
            Expression::OrElse(left, right) => {
                // Short-circuit: if left is true, don't evaluate right
                let l = self.evaluate_expr(left)?;
                if l.as_bool()? {
                    return Ok(Value::Boolean(true));
                }
                let r = self.evaluate_expr(right)?;
                Ok(Value::Boolean(r.as_bool()?))
            }
            Expression::Is(left, right) => {
                // Reference equality — for Nothing comparison and object identity
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let result = match (&l, &r) {
                    (Value::Nothing, Value::Nothing) => true,
                    (Value::Nothing, _) | (_, Value::Nothing) => {
                        matches!((&l, &r), (Value::Nothing, Value::Nothing))
                    }
                    (Value::Object(a), Value::Object(b)) => std::ptr::eq(&*a.borrow(), &*b.borrow()),
                    _ => l == r,
                };
                Ok(Value::Boolean(result))
            }
            Expression::IsNot(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let result = match (&l, &r) {
                    (Value::Nothing, Value::Nothing) => false,
                    (Value::Nothing, _) | (_, Value::Nothing) => true,
                    (Value::Object(a), Value::Object(b)) => !std::ptr::eq(&*a.borrow(), &*b.borrow()),
                    _ => l != r,
                };
                Ok(Value::Boolean(result))
            }
            Expression::Like(left, right) => {
                // VB.NET Like operator — basic pattern matching with *, ?, #, [charlist]
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let text = l.as_string();
                let pattern = r.as_string();
                let result = vb_like_match(&text, &pattern);
                Ok(Value::Boolean(result))
            }
            Expression::TypeOf { expr, type_name } => {
                let val = self.evaluate_expr(expr)?;
                let tn = type_name.trim();
                let result = match &val {
                    Value::Object(obj) => {
                        let b = obj.borrow();
                        self.is_type_or_base(&b.class_name, tn)
                    }
                    Value::String(_) => tn.eq_ignore_ascii_case("String"),
                    Value::Integer(_) => tn.eq_ignore_ascii_case("Integer") || tn.eq_ignore_ascii_case("Int32"),
                    Value::Long(_) => tn.eq_ignore_ascii_case("Long") || tn.eq_ignore_ascii_case("Int64"),
                    Value::Double(_) => tn.eq_ignore_ascii_case("Double"),
                    Value::Boolean(_) => tn.eq_ignore_ascii_case("Boolean"),
                    Value::Nothing => false,
                    _ => false,
                };
                Ok(Value::Boolean(result))
            }
            Expression::BitShiftLeft(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let val = l.as_long()?;
                let shift = r.as_integer()?;
                // Rust panics if shift > 64, need to handle?
                // VB wraps or uses mask? VB: "The shift amount is masked to the size of the data type"
                // Long: count And 63. Integer: count And 31.
                // We treat everything as Long here basically.
                let shift_masked = shift & 63;
                Ok(Value::Long(val << shift_masked))
            }
            Expression::BitShiftRight(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let val = l.as_long()?;
                let shift = r.as_integer()?;
                let shift_masked = shift & 63;
                Ok(Value::Long(val >> shift_masked))
            }
            Expression::Modulo(left, right) => {
                let l = self.evaluate_expr(left)?;
                let r = self.evaluate_expr(right)?;
                let a = l.as_integer()?;
                let b = r.as_integer()?;
                if b == 0 { return Err(RuntimeError::DivisionByZero); }
                Ok(Value::Integer(a % b))
            }
            Expression::Me => {
                if let Some(obj_rc) = &self.current_object {
                    Ok(Value::Object(obj_rc.clone()))
                } else {
                    Err(RuntimeError::Custom("'Me' used outside of object context".to_string()))
                }
            }
            Expression::MyBase => {
                // MyBase refers to the same object as Me, but method
                // dispatch starts from the parent class.  We return the
                // same Rc; call_method checks for Expression::MyBase to
                // resolve the method from the base class.
                if let Some(obj_rc) = &self.current_object {
                    Ok(Value::Object(obj_rc.clone()))
                } else {
                    Err(RuntimeError::Custom("'MyBase' used outside of object context".to_string()))
                }
            }
            Expression::WithTarget => {
                if let Some(val) = &self.with_object {
                    Ok(val.clone())
                } else {
                    Err(RuntimeError::Custom("'.' used outside of With block".to_string()))
                }
            }
            Expression::Cast { expr, .. } => {
                // CType/DirectCast/TryCast — in our dynamically typed interpreter,
                // just evaluate the inner expression (the cast is a no-op at runtime).
                self.evaluate_expr(expr)
            }
            Expression::IfExpression(first, second, third) => {
                if let Some(false_expr) = third {
                    // Ternary: If(condition, trueValue, falseValue)
                    let cond = self.evaluate_expr(first)?;
                    if cond.is_truthy() {
                        self.evaluate_expr(second)
                    } else {
                        self.evaluate_expr(false_expr)
                    }
                } else {
                    // Coalesce: If(value, default) — return value if not Nothing, else default
                    let val = self.evaluate_expr(first)?;
                    match &val {
                        Value::Nothing => self.evaluate_expr(second),
                        _ => Ok(val),
                    }
                }
            }
            Expression::AddressOf(name) => {
                // For now, store as a string value — delegates are not fully supported
                Ok(Value::String(format!("AddressOf:{}", name)))
            }
            Expression::New(class_id, ctor_args) => {
                // Strip generic suffix: "List(Of String)" → "list"
                let class_name_full = class_id.as_str().to_lowercase();
                let class_name = if let Some(idx) = class_name_full.find("(of ") {
                    class_name_full[..idx].trim_end().to_string()
                } else {
                    class_name_full
                };

                // ===== XML CONSTRUCTORS: XDocument, XElement, XAttribute =====
                if class_name == "xelement" || class_name == "system.xml.linq.xelement" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    return Ok(crate::builtins::xml::create_xelement(&arg_values));
                }
                if class_name == "xattribute" || class_name == "system.xml.linq.xattribute" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    return Ok(crate::builtins::xml::create_xattribute(&arg_values));
                }
                if class_name == "xdocument" || class_name == "system.xml.linq.xdocument" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    return Ok(crate::builtins::xml::create_xdocument(&arg_values));
                }
                if class_name == "xcomment" || class_name == "system.xml.linq.xcomment" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    return Ok(crate::builtins::xml::create_xcomment(&arg_values));
                }
                if class_name == "xdeclaration" || class_name == "system.xml.linq.xdeclaration" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    return Ok(crate::builtins::xml::create_xdeclaration(&arg_values));
                }

                // Handle Common Dialogs
                if class_name == "openfiledialog" || class_name.ends_with(".openfiledialog") {
                    return Ok(crate::builtins::dialogs::create_openfiledialog());
                }
                if class_name == "savefiledialog" || class_name.ends_with(".savefiledialog") {
                    return Ok(crate::builtins::dialogs::create_savefiledialog());
                }
                if class_name == "colordialog" || class_name.ends_with(".colordialog") {
                    return Ok(crate::builtins::dialogs::create_colordialog());
                }
                if class_name == "fontdialog" || class_name.ends_with(".fontdialog") {
                    return Ok(crate::builtins::dialogs::create_fontdialog());
                }
                if class_name == "folderbrowserdialog" || class_name.ends_with(".folderbrowserdialog") {
                    return Ok(crate::builtins::dialogs::create_folderbrowserdialog());
                }
                if class_name == "printdialog" || class_name.ends_with(".printdialog") {
                    return Ok(crate::builtins::dialogs::create_printdialog());
                }
                if class_name == "printpreviewdialog" || class_name.ends_with(".printpreviewdialog") {
                    return Ok(crate::builtins::dialogs::create_printpreviewdialog());
                }

                // ===== System.Drawing =====
                if class_name == "pen" || class_name == "system.drawing.pen" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    return crate::builtins::drawing_fns::new_pen_fn(&arg_values);
                }
                if class_name == "solidbrush" || class_name == "system.drawing.solidbrush" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    return crate::builtins::drawing_fns::new_solid_brush_fn(&arg_values);
                }


                if class_name == "pagesetupdialog" || class_name.ends_with(".pagesetupdialog") {
                    return Ok(crate::builtins::dialogs::create_pagesetupdialog());
                }
                if class_name == "helpprovider" || class_name.ends_with(".helpprovider") {
                    return Ok(crate::builtins::dialogs::create_helpprovider());
                }

                // ===== EVENTARGS TYPE CONSTRUCTORS =====
                {
                    let ea_lower = class_name.to_lowercase();
                    match ea_lower.as_str() {
                        "eventargs" | "system.eventargs" => {
                            return Ok(Self::make_event_args());
                        }
                        "mouseeventargs" | "system.windows.forms.mouseeventargs" => {
                            let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                            let arg_values = arg_values?;
                            let button = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                            let clicks = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(1);
                            let x = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                            let y = arg_values.get(3).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                            let delta = arg_values.get(4).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                            return Ok(Self::make_mouse_event_args(button, clicks, x, y, delta));
                        }
                        "keyeventargs" | "system.windows.forms.keyeventargs" => {
                            let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                            let arg_values = arg_values?;
                            let key_code = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                            return Ok(Self::make_key_event_args(key_code, false, false, false));
                        }
                        "keypresseventargs" | "system.windows.forms.keypresseventargs" => {
                            let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                            let arg_values = arg_values?;
                            let key_char = arg_values.get(0)
                                .and_then(|v| if let Value::Char(c) = v { Some(*c) } else { None })
                                .unwrap_or('\0');
                            return Ok(Self::make_key_press_event_args(key_char));
                        }
                        "formclosingeventargs" | "system.windows.forms.formclosingeventargs" => {
                            let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                            let arg_values = arg_values?;
                            let reason = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(3);
                            return Ok(Self::make_form_closing_event_args(reason));
                        }
                        "formclosedeventargs" | "system.windows.forms.formclosedeventargs" => {
                            let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                            let arg_values = arg_values?;
                            let reason = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(3);
                            return Ok(Self::make_form_closed_event_args(reason));
                        }
                        "painteventargs" | "system.windows.forms.painteventargs" => {
                            return Ok(Self::make_paint_event_args());
                        }
                        "processstartinfo" | "system.diagnostics.processstartinfo" => {
                            let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                            let arg_values = arg_values?;
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("filename".to_string(), arg_values.get(0).cloned().unwrap_or(Value::String(String::new())));
                            fields.insert("arguments".to_string(), arg_values.get(1).cloned().unwrap_or(Value::String(String::new())));
                            fields.insert("__type".to_string(), Value::String("ProcessStartInfo".to_string()));
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(crate::value::ObjectData {
                                class_name: "ProcessStartInfo".to_string(),
                                fields,
                            }))));
                        }
                        _ => {}
                    }
                }

                // ===== EXCEPTION TYPE CONSTRUCTORS =====
                {
                    let exception_types = [
                        "exception", "system.exception",
                        "argumentexception", "system.argumentexception",
                        "argumentnullexception", "system.argumentnullexception",
                        "argumentoutofrangeexception", "system.argumentoutofrangeexception",
                        "invalidoperationexception", "system.invalidoperationexception",
                        "invalidcastexception", "system.invalidcastexception",
                        "notsupportedexception", "system.notsupportedexception",
                        "notimplementedexception", "system.notimplementedexception",
                        "nullreferenceexception", "system.nullreferenceexception",
                        "indexoutofrangeexception", "system.indexoutofrangeexception",
                        "keynotfoundexception", "system.collections.generic.keynotfoundexception",
                        "filenotfoundexception", "system.io.filenotfoundexception",
                        "directorynotfoundexception", "system.io.directorynotfoundexception",
                        "ioexception", "system.io.ioexception",
                        "formatexception", "system.formatexception",
                        "overflowexception", "system.overflowexception",
                        "dividebyzeroexception", "system.dividebyzeroexception",
                        "stackoverflowexception", "system.stackoverflowexception",
                        "outofmemoryexception", "system.outofmemoryexception",
                        "timeoutexception", "system.timeoutexception",
                        "operationcanceledexception", "system.operationcanceledexception",
                        "unauthorizedaccessexception", "system.unauthorizedaccessexception",
                        "applicationexception", "system.applicationexception",
                        "aggregateexception", "system.aggregateexception",
                        "taskcanceledexception", "system.threading.tasks.taskcanceledexception",
                        "objectdisposedexception", "system.objectdisposedexception",
                        "socketsexception", "system.net.sockets.socketexception",
                    ];
                    if exception_types.contains(&class_name.as_str()) {
                        let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                        let arg_values = arg_values?;
                        let nice_name = class_id.as_str().split('.').last().unwrap_or(class_id.as_str()).to_string();
                        let msg = arg_values.get(0).map(|v| v.as_string()).unwrap_or_else(|| format!("Exception of type '{}' was thrown.", nice_name));
                        let inner = arg_values.get(1).and_then(|v| {
                            if *v == Value::Nothing { None } else { Some(v.as_string()) }
                        });
                        let mut fields = std::collections::HashMap::new();
                        fields.insert("message".to_string(), Value::String(msg));
                        fields.insert("stacktrace".to_string(), Value::String(String::new()));
                        fields.insert("source".to_string(), Value::String(String::new()));
                        fields.insert("hresult".to_string(), Value::Integer(-2146233088));
                        fields.insert("innerexception".to_string(), inner.map(Value::String).unwrap_or(Value::Nothing));
                        fields.insert("__type".to_string(), Value::String(nice_name.clone()));
                        let obj = crate::value::ObjectData { class_name: nice_name, fields };
                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                    }
                }

                // ===== TUPLE / VALUETUPLE CONSTRUCTORS =====
                if class_name == "tuple" || class_name == "system.tuple"
                    || class_name == "valuetuple" || class_name == "system.valuetuple" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Tuple".to_string()));
                    for (i, val) in arg_values.iter().enumerate() {
                        fields.insert(format!("item{}", i + 1), val.clone());
                    }
                    fields.insert("length".to_string(), Value::Integer(arg_values.len() as i32));
                    let obj = crate::value::ObjectData { class_name: "Tuple".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== NULLABLE(OF T) =====
                if class_name.starts_with("nullable") || class_name.starts_with("system.nullable") {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let inner_val = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                    let has_value = inner_val != Value::Nothing;
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Nullable".to_string()));
                    fields.insert("value".to_string(), inner_val);
                    fields.insert("hasvalue".to_string(), Value::Boolean(has_value));
                    let obj = crate::value::ObjectData { class_name: "Nullable".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.THREADING.THREAD =====
                if class_name == "thread" || class_name == "system.threading.thread" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Thread".to_string()));
                    if let Some(task) = arg_values.get(0) {
                        fields.insert("__task".to_string(), task.clone());
                    }
                    fields.insert("isalive".to_string(), Value::Boolean(false));
                    fields.insert("managedthreadid".to_string(), Value::Integer(0));
                    let obj = crate::value::ObjectData { class_name: "Thread".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.TIMERS.TIMER =====
                if class_name == "timer" || class_name == "system.timers.timer" || class_name == "system.threading.timer" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let interval = arg_values.get(0).and_then(|v| v.as_double().ok()).unwrap_or(100.0);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Timer".to_string()));
                    fields.insert("interval".to_string(), Value::Double(interval));
                    fields.insert("enabled".to_string(), Value::Boolean(false));
                    fields.insert("autoreset".to_string(), Value::Boolean(true));
                    fields.insert("__elapsed_count".to_string(), Value::Integer(0));
                    let obj = crate::value::ObjectData { class_name: "Timer".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.IO.FILESTREAM =====
                if class_name == "filestream" || class_name == "system.io.filestream" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let path = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                    let mode = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(3); // 3=OpenOrCreate
                    let access = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(3); // 3=ReadWrite
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("FileStream".to_string()));
                    fields.insert("__path".to_string(), Value::String(path.clone()));
                    fields.insert("__mode".to_string(), Value::Integer(mode));
                    fields.insert("__access".to_string(), Value::Integer(access));
                    fields.insert("__position".to_string(), Value::Long(0));
                    fields.insert("name".to_string(), Value::String(path.clone()));
                    fields.insert("canread".to_string(), Value::Boolean(access == 1 || access == 3));
                    fields.insert("canwrite".to_string(), Value::Boolean(access == 2 || access == 3));
                    fields.insert("canseek".to_string(), Value::Boolean(true));
                    // Read file contents into buffer
                    let data = if std::path::Path::new(&path).exists() && (mode != 2) {
                        std::fs::read(&path).unwrap_or_default()
                    } else {
                        Vec::new()
                    };
                    fields.insert("length".to_string(), Value::Long(data.len() as i64));
                    fields.insert("__data".to_string(), Value::Array(data.iter().map(|b| Value::Integer(*b as i32)).collect()));
                    fields.insert("position".to_string(), Value::Long(0));
                    fields.insert("__closed".to_string(), Value::Boolean(false));
                    let obj = crate::value::ObjectData { class_name: "FileStream".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.IO.MEMORYSTREAM =====
                if class_name == "memorystream" || class_name == "system.io.memorystream" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let initial_data = if let Some(Value::Array(arr)) = arg_values.get(0) {
                        arr.clone()
                    } else if let Some(val) = arg_values.get(0) {
                        let cap = val.as_integer().unwrap_or(0);
                        vec![Value::Integer(0); cap.max(0) as usize]
                    } else {
                        Vec::new()
                    };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("MemoryStream".to_string()));
                    fields.insert("__data".to_string(), Value::Array(initial_data.clone()));
                    fields.insert("length".to_string(), Value::Long(initial_data.len() as i64));
                    fields.insert("position".to_string(), Value::Long(0));
                    fields.insert("capacity".to_string(), Value::Long(initial_data.len() as i64));
                    fields.insert("canread".to_string(), Value::Boolean(true));
                    fields.insert("canwrite".to_string(), Value::Boolean(true));
                    fields.insert("canseek".to_string(), Value::Boolean(true));
                    fields.insert("__closed".to_string(), Value::Boolean(false));
                    let obj = crate::value::ObjectData { class_name: "MemoryStream".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.IO.BINARYREADER =====
                if class_name == "binaryreader" || class_name == "system.io.binaryreader" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    // First arg is a stream (FileStream or MemoryStream object)
                    let stream_data = if let Some(Value::Object(sref)) = arg_values.get(0) {
                        let sb = sref.borrow();
                        sb.fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()))
                    } else {
                        Value::Array(Vec::new())
                    };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("BinaryReader".to_string()));
                    fields.insert("__data".to_string(), stream_data);
                    fields.insert("__position".to_string(), Value::Long(0));
                    fields.insert("__closed".to_string(), Value::Boolean(false));
                    let obj = crate::value::ObjectData { class_name: "BinaryReader".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.IO.BINARYWRITER =====
                if class_name == "binarywriter" || class_name == "system.io.binarywriter" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let stream_ref = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("BinaryWriter".to_string()));
                    fields.insert("__stream".to_string(), stream_ref);
                    fields.insert("__data".to_string(), Value::Array(Vec::new()));
                    fields.insert("__closed".to_string(), Value::Boolean(false));
                    let obj = crate::value::ObjectData { class_name: "BinaryWriter".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.NET.SOCKETS.TCPCLIENT =====
                if class_name == "tcpclient" || class_name == "system.net.sockets.tcpclient" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let host = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                    let port = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("TcpClient".to_string()));
                    fields.insert("__host".to_string(), Value::String(host.clone()));
                    fields.insert("__port".to_string(), Value::Integer(port));
                    fields.insert("receivebuffersize".to_string(), Value::Integer(8192));
                    fields.insert("sendbuffersize".to_string(), Value::Integer(8192));
                    // If host and port provided, actually connect
                    if !host.is_empty() && port > 0 {
                        match crate::builtins::networking::NetHandle::connect_tcp(&host, port, None) {
                            Ok(handle) => {
                                let id = self.alloc_net_handle(handle);
                                fields.insert("__socket_id".to_string(), Value::Long(id));
                                fields.insert("connected".to_string(), Value::Boolean(true));
                            }
                            Err(e) => {
                                return Err(RuntimeError::Custom(format!("TcpClient connect failed: {}", e)));
                            }
                        }
                    } else {
                        fields.insert("__socket_id".to_string(), Value::Long(0));
                        fields.insert("connected".to_string(), Value::Boolean(false));
                    }
                    let obj = crate::value::ObjectData { class_name: "TcpClient".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.NET.SOCKETS.TCPLISTENER =====
                if class_name == "tcplistener" || class_name == "system.net.sockets.tcplistener" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    // TcpListener(IPAddress, port) or TcpListener(port)
                    let (addr, port) = if arg_values.len() >= 2 {
                        (arg_values[0].as_string(), arg_values[1].as_integer().unwrap_or(0))
                    } else {
                        ("0.0.0.0".to_string(), arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0))
                    };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("TcpListener".to_string()));
                    fields.insert("__address".to_string(), Value::String(addr));
                    fields.insert("__port".to_string(), Value::Integer(port));
                    fields.insert("__active".to_string(), Value::Boolean(false));
                    fields.insert("__socket_id".to_string(), Value::Long(0));
                    let obj = crate::value::ObjectData { class_name: "TcpListener".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.NET.SOCKETS.UDPCLIENT =====
                if class_name == "udpclient" || class_name == "system.net.sockets.udpclient" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let port = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("UdpClient".to_string()));
                    fields.insert("__port".to_string(), Value::Integer(port));
                    // Actually bind the UDP socket
                    match crate::builtins::networking::NetHandle::bind_udp(port) {
                        Ok(handle) => {
                            let id = self.alloc_net_handle(handle);
                            fields.insert("__socket_id".to_string(), Value::Long(id));
                        }
                        Err(e) => {
                            return Err(RuntimeError::Custom(format!("UdpClient bind failed: {}", e)));
                        }
                    }
                    let obj = crate::value::ObjectData { class_name: "UdpClient".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.NET.MAIL.SMTPCLIENT =====
                if class_name == "smtpclient" || class_name == "system.net.mail.smtpclient" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let host = arg_values.get(0).map(|v| v.as_string()).unwrap_or_else(|| "localhost".to_string());
                    let port = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(25);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("SmtpClient".to_string()));
                    fields.insert("host".to_string(), Value::String(host));
                    fields.insert("port".to_string(), Value::Integer(port));
                    fields.insert("enablessl".to_string(), Value::Boolean(false));
                    fields.insert("credentials".to_string(), Value::Nothing);
                    fields.insert("deliverymethod".to_string(), Value::Integer(0)); // Network
                    let obj = crate::value::ObjectData { class_name: "SmtpClient".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.NET.MAIL.MAILMESSAGE =====
                if class_name == "mailmessage" || class_name == "system.net.mail.mailmessage" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let from = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                    let to = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                    let subject = arg_values.get(2).map(|v| v.as_string()).unwrap_or_default();
                    let body = arg_values.get(3).map(|v| v.as_string()).unwrap_or_default();
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("MailMessage".to_string()));
                    fields.insert("from".to_string(), Value::String(from));
                    fields.insert("to".to_string(), Value::String(to));
                    fields.insert("subject".to_string(), Value::String(subject));
                    fields.insert("body".to_string(), Value::String(body));
                    fields.insert("isbodyhtml".to_string(), Value::Boolean(false));
                    fields.insert("cc".to_string(), Value::String(String::new()));
                    fields.insert("bcc".to_string(), Value::String(String::new()));
                    let obj = crate::value::ObjectData { class_name: "MailMessage".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.NET.MAIL.MAILADDRESS =====
                if class_name == "mailaddress" || class_name == "system.net.mail.mailaddress" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let address = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                    let display_name = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("MailAddress".to_string()));
                    fields.insert("address".to_string(), Value::String(address.clone()));
                    fields.insert("displayname".to_string(), Value::String(display_name));
                    fields.insert("host".to_string(), Value::String(address.split('@').nth(1).unwrap_or("").to_string()));
                    fields.insert("user".to_string(), Value::String(address.split('@').next().unwrap_or("").to_string()));
                    let obj = crate::value::ObjectData { class_name: "MailAddress".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.THREADING.MUTEX =====
                if class_name == "mutex" || class_name == "system.threading.mutex" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let initially_owned = arg_values.get(0).and_then(|v| v.as_bool().ok()).unwrap_or(false);
                    let name = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Mutex".to_string()));
                    fields.insert("__owned".to_string(), Value::Boolean(initially_owned));
                    fields.insert("__name".to_string(), Value::String(name));
                    let obj = crate::value::ObjectData { class_name: "Mutex".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.THREADING.SEMAPHORE =====
                if class_name == "semaphore" || class_name == "system.threading.semaphore"
                    || class_name == "semaphoreslim" || class_name == "system.threading.semaphoreslim" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let initial_count = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(1);
                    let max_count = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(initial_count);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Semaphore".to_string()));
                    fields.insert("__count".to_string(), Value::Integer(initial_count));
                    fields.insert("__max".to_string(), Value::Integer(max_count));
                    fields.insert("currentcount".to_string(), Value::Integer(initial_count));
                    let obj = crate::value::ObjectData { class_name: "Semaphore".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.DRAWING.POINT =====
                if class_name == "point" || class_name == "system.drawing.point" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let x = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let y = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Point".to_string()));
                    fields.insert("x".to_string(), Value::Integer(x));
                    fields.insert("y".to_string(), Value::Integer(y));
                    fields.insert("isempty".to_string(), Value::Boolean(x == 0 && y == 0));
                    let obj = crate::value::ObjectData { class_name: "Point".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.DRAWING.SIZE =====
                if class_name == "size" || class_name == "system.drawing.size" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let w = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let h = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Size".to_string()));
                    fields.insert("width".to_string(), Value::Integer(w));
                    fields.insert("height".to_string(), Value::Integer(h));
                    fields.insert("isempty".to_string(), Value::Boolean(w == 0 && h == 0));
                    let obj = crate::value::ObjectData { class_name: "Size".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.DRAWING.RECTANGLE =====
                if class_name == "rectangle" || class_name == "system.drawing.rectangle" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let x = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let y = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let w = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let h = arg_values.get(3).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Rectangle".to_string()));
                    fields.insert("x".to_string(), Value::Integer(x));
                    fields.insert("y".to_string(), Value::Integer(y));
                    fields.insert("width".to_string(), Value::Integer(w));
                    fields.insert("height".to_string(), Value::Integer(h));
                    fields.insert("left".to_string(), Value::Integer(x));
                    fields.insert("top".to_string(), Value::Integer(y));
                    fields.insert("right".to_string(), Value::Integer(x + w));
                    fields.insert("bottom".to_string(), Value::Integer(y + h));
                    let obj = crate::value::ObjectData { class_name: "Rectangle".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.DRAWING.FONT =====
                if class_name == "font" || class_name == "system.drawing.font" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let family = arg_values.get(0).map(|v| v.as_string()).unwrap_or("Microsoft Sans Serif".to_string());
                    let size = arg_values.get(1).map(|v| match v { Value::Double(d) => *d, Value::Single(f) => *f as f64, Value::Integer(i) => *i as f64, _ => 8.25 }).unwrap_or(8.25);
                    let style = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(0); // 0=Regular
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Font".to_string()));
                    fields.insert("name".to_string(), Value::String(family.clone()));
                    fields.insert("fontfamily".to_string(), Value::String(family));
                    fields.insert("size".to_string(), Value::Double(size));
                    fields.insert("sizeininpoints".to_string(), Value::Double(size));
                    fields.insert("bold".to_string(), Value::Boolean(style & 1 != 0));
                    fields.insert("italic".to_string(), Value::Boolean(style & 2 != 0));
                    fields.insert("underline".to_string(), Value::Boolean(style & 4 != 0));
                    fields.insert("strikeout".to_string(), Value::Boolean(style & 8 != 0));
                    fields.insert("style".to_string(), Value::Integer(style));
                    let obj = crate::value::ObjectData { class_name: "Font".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== SYSTEM.DRAWING.COLOR =====
                if class_name == "color" || class_name == "system.drawing.color" {
                    // Color.FromArgb is handled separately; New Color() returns Empty
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Color".to_string()));
                    fields.insert("r".to_string(), Value::Integer(0));
                    fields.insert("g".to_string(), Value::Integer(0));
                    fields.insert("b".to_string(), Value::Integer(0));
                    fields.insert("a".to_string(), Value::Integer(255));
                    fields.insert("name".to_string(), Value::String("Empty".to_string()));
                    fields.insert("isempty".to_string(), Value::Boolean(true));
                    let obj = crate::value::ObjectData { class_name: "Color".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== STRINGBUILDER =====
                if class_name == "stringbuilder" || class_name == "system.text.stringbuilder" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    return crate::builtins::text_fns::stringbuilder_new_fn(&arg_values);
                }

                // ===== PROCESS =====
                if class_name.eq_ignore_ascii_case("process") || class_name.eq_ignore_ascii_case("system.diagnostics.process") {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Process".to_string()));
                    
                    // StartInfo property: ProcessStartInfo object
                    let mut si_fields = std::collections::HashMap::new();
                    si_fields.insert("__type".to_string(), Value::String("ProcessStartInfo".to_string()));
                    si_fields.insert("filename".to_string(), Value::String(String::new()));
                    si_fields.insert("arguments".to_string(), Value::String(String::new()));
                    si_fields.insert("useshellexecute".to_string(), Value::Boolean(true));
                    let start_info = crate::value::ObjectData { class_name: "ProcessStartInfo".to_string(), fields: si_fields };
                    let si_obj = Value::Object(std::rc::Rc::new(std::cell::RefCell::new(start_info)));
                    
                    fields.insert("startinfo".to_string(), si_obj);
                    
                    let obj = crate::value::ObjectData { class_name: "Process".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // Check if this is a short WinForms control name
                let is_short_winforms = matches!(class_name.as_str(),
                    "datetimepicker" | "linklabel" | "toolstrip" | "trackbar" |
                    "maskedtextbox" | "splitcontainer" | "flowlayoutpanel" | "tablelayoutpanel" |
                    "monthcalendar" | "hscrollbar" | "vscrollbar" | "tooltip" |
                    "textbox" | "label" | "button" | "checkbox" | "radiobutton" |
                    "groupbox" | "panel" | "combobox" | "listbox" | "picturebox" |
                    "richtextbox" | "webbrowser" | "treeview" | "listview" |
                    "datagridview" | "tabcontrol" | "tabpage" | "progressbar" |
                    "numericupdown" | "menustrip" | "contextmenustrip" | "statusstrip" |
                    "toolstripmenuitem" | "toolstripstatuslabel"
                );

                if is_short_winforms || (class_name.starts_with("system.windows.forms.")
                    && class_name != "system.windows.forms.bindingsource")
                {
                    // Create a proper control object with fields that mimic WinForms properties
                    let base_name = class_id.as_str().split('.').last().unwrap_or(class_id.as_str()).to_string();
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String(base_name.clone()));
                    fields.insert("__is_control".to_string(), Value::Boolean(true));
                    fields.insert("name".to_string(), Value::String(String::new()));
                    fields.insert("text".to_string(), Value::String(String::new()));
                    fields.insert("visible".to_string(), Value::Boolean(true));
                    fields.insert("enabled".to_string(), Value::Boolean(true));
                    fields.insert("left".to_string(), Value::Integer(0));
                    fields.insert("top".to_string(), Value::Integer(0));
                    fields.insert("width".to_string(), Value::Integer(100));
                    fields.insert("height".to_string(), Value::Integer(30));
                    fields.insert("tag".to_string(), Value::Nothing);
                    fields.insert("tabindex".to_string(), Value::Integer(0));
                    fields.insert("backcolor".to_string(), Value::String(String::new()));
                    fields.insert("forecolor".to_string(), Value::String(String::new()));
                    fields.insert("font".to_string(), Value::Nothing);
                    fields.insert("anchor".to_string(), Value::Integer(5)); // Top | Left
                    fields.insert("dock".to_string(), Value::Integer(0)); // None

                    // Type-specific default fields
                    Self::init_control_type_defaults(&base_name, &mut fields);

                    let obj = crate::value::ObjectData { class_name: base_name, fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                if class_name.starts_with("system.drawing.") 
                    || class_name.starts_with("system.componentmodel.") {
                    return Ok(Value::Nothing);
                }

                if class_name == "system.collections.arraylist" || class_name == "arraylist"
                    || class_name == "collection" || class_name == "system.collections.collection" {
                    return Ok(Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                }
                
                // Generic List (parsed as "List" or "System.Collections.Generic.List")
                if class_name.ends_with(".list") || class_name == "list" {
                     return Ok(Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))));
                }

                // Queue
                if class_name == "queue" || class_name == "system.collections.queue"
                    || class_name == "system.collections.generic.queue" {
                    return Ok(Value::Queue(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::Queue::new()))));
                }

                // Stack
                if class_name == "stack" || class_name == "system.collections.stack"
                    || class_name == "system.collections.generic.stack" {
                    return Ok(Value::Stack(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::Stack::new()))));
                }

                // HashSet
                if class_name == "hashset" || class_name == "system.collections.generic.hashset" {
                    return Ok(Value::HashSet(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::VBHashSet::new()))));
                }

                // Dictionary
                if class_name == "dictionary" || class_name == "system.collections.generic.dictionary" {
                     return Ok(Value::Dictionary(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::VBDictionary::new()))));
                }

                // ConcurrentDictionary
                if class_name == "concurrentdictionary" || class_name == "system.collections.concurrent.concurrentdictionary" {
                     return Ok(Value::ConcurrentDictionary(crate::builtins::concurrent_collections::ConcurrentDictionary::new()));
                }

                // ConcurrentQueue
                if class_name == "concurrentqueue" || class_name == "system.collections.concurrent.concurrentqueue" {
                     return Ok(Value::ConcurrentQueue(crate::builtins::concurrent_collections::ConcurrentQueue::new()));
                }

                // ConcurrentStack
                if class_name == "concurrentstack" || class_name == "system.collections.concurrent.concurrentstack" {
                     return Ok(Value::ConcurrentStack(crate::builtins::concurrent_collections::ConcurrentStack::new()));
                }
                if class_name == "dictionary" || class_name == "system.collections.generic.dictionary" {
                     return Ok(Value::Dictionary(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::VBDictionary::new()))));
                }

                // ConcurrentDictionary
                if class_name == "concurrentdictionary" || class_name == "system.collections.concurrent.concurrentdictionary" {
                     return Ok(Value::ConcurrentDictionary(crate::builtins::concurrent_collections::ConcurrentDictionary::new()));
                }

                // ConcurrentQueue
                if class_name == "concurrentqueue" || class_name == "system.collections.concurrent.concurrentqueue" {
                     return Ok(Value::ConcurrentQueue(crate::builtins::concurrent_collections::ConcurrentQueue::new()));
                }

                // ConcurrentStack
                if class_name == "concurrentstack" || class_name == "system.collections.concurrent.concurrentstack" {
                     return Ok(Value::ConcurrentStack(crate::builtins::concurrent_collections::ConcurrentStack::new()));
                }
                if class_name == "dictionary" || class_name == "system.collections.generic.dictionary"
                    || class_name == "system.collections.hashtable" || class_name == "hashtable" {
                    return Ok(Value::Dictionary(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::VBDictionary::new()))));
                }

                // WebClient (System.Net.WebClient) – backed by curl
                if class_name == "webclient" || class_name == "system.net.webclient" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("WebClient".to_string()));
                    fields.insert("Encoding".to_string(), Value::String("UTF-8".to_string()));
                    // Headers stored as a Dictionary for curl -H flags
                    fields.insert("headers".to_string(), Value::Dictionary(
                        std::rc::Rc::new(std::cell::RefCell::new(crate::collections::VBDictionary::new()))
                    ));
                    let obj = crate::value::ObjectData {
                        class_name: "WebClient".to_string(),
                        fields,
                    };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // HttpClient (System.Net.Http.HttpClient) – backed by curl
                if class_name == "httpclient" || class_name == "system.net.http.httpclient" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("HttpClient".to_string()));
                    // DefaultRequestHeaders stored as a Dictionary
                    fields.insert("defaultrequestheaders".to_string(), Value::Dictionary(
                        std::rc::Rc::new(std::cell::RefCell::new(crate::collections::VBDictionary::new()))
                    ));
                    let obj = crate::value::ObjectData {
                        class_name: "HttpClient".to_string(),
                        fields,
                    };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // NetworkCredential
                if class_name == "networkcredential" || class_name == "system.net.networkcredential" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("NetworkCredential".to_string()));
                    if ctor_args.len() >= 2 {
                        let user = self.evaluate_expr(&ctor_args[0])?.as_string();
                        let pass = self.evaluate_expr(&ctor_args[1])?.as_string();
                        fields.insert("username".to_string(), Value::String(user));
                        fields.insert("password".to_string(), Value::String(pass));
                        if ctor_args.len() >= 3 {
                            let domain = self.evaluate_expr(&ctor_args[2])?.as_string();
                            fields.insert("domain".to_string(), Value::String(domain));
                        }
                    }
                    let obj = crate::value::ObjectData {
                        class_name: "NetworkCredential".to_string(),
                        fields,
                    };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // System.DateTime
                if class_name == "datetime" || class_name == "system.datetime" || class_name == "date" {
                    // New DateTime(year, month, day) or New DateTime(year, month, day, hour, minute, second)
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let year = arg_values.get(0).map(|v| v.as_integer().unwrap_or(1)).unwrap_or(1) as i32;
                    let month = arg_values.get(1).map(|v| v.as_integer().unwrap_or(1)).unwrap_or(1) as u32;
                    let day = arg_values.get(2).map(|v| v.as_integer().unwrap_or(1)).unwrap_or(1) as u32;
                    let hour = arg_values.get(3).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0) as u32;
                    let minute = arg_values.get(4).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0) as u32;
                    let second = arg_values.get(5).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0) as u32;
                    let dt = chrono::NaiveDate::from_ymd_opt(year, month, day)
                        .and_then(|d| d.and_hms_opt(hour, minute, second))
                        .unwrap_or_else(|| chrono::NaiveDate::from_ymd_opt(2000, 1, 1).unwrap().and_hms_opt(0, 0, 0).unwrap());
                    return Ok(Value::Date(date_to_ole(dt)));
                }

                // New TimeSpan(hours, minutes, seconds) or New TimeSpan(days, hours, minutes, seconds, ms)
                if class_name == "timespan" || class_name == "system.timespan" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let (days, hours, minutes, seconds, ms) = match arg_values.len() {
                        1 => {
                            // New TimeSpan(ticks)
                            let ticks = arg_values[0].as_double().unwrap_or(0.0);
                            let total_ms = ticks / 10000.0;
                            (0, 0, 0, 0, total_ms as i32)
                        }
                        3 => {
                            // New TimeSpan(hours, minutes, seconds)
                            let h = arg_values[0].as_integer().unwrap_or(0);
                            let m = arg_values[1].as_integer().unwrap_or(0);
                            let s = arg_values[2].as_integer().unwrap_or(0);
                            (0, h, m, s, 0)
                        }
                        4 => {
                            // New TimeSpan(days, hours, minutes, seconds)
                            let d = arg_values[0].as_integer().unwrap_or(0);
                            let h = arg_values[1].as_integer().unwrap_or(0);
                            let m = arg_values[2].as_integer().unwrap_or(0);
                            let s = arg_values[3].as_integer().unwrap_or(0);
                            (d, h, m, s, 0)
                        }
                        5 => {
                            // New TimeSpan(days, hours, minutes, seconds, ms)
                            let d = arg_values[0].as_integer().unwrap_or(0);
                            let h = arg_values[1].as_integer().unwrap_or(0);
                            let m = arg_values[2].as_integer().unwrap_or(0);
                            let s = arg_values[3].as_integer().unwrap_or(0);
                            let ms = arg_values[4].as_integer().unwrap_or(0);
                            (d, h, m, s, ms)
                        }
                        _ => (0, 0, 0, 0, 0),
                    };
                    let total_seconds = (days as f64) * 86400.0 + (hours as f64) * 3600.0 + (minutes as f64) * 60.0 + (seconds as f64) + (ms as f64) / 1000.0;
                    let obj_data = crate::value::ObjectData {
                        class_name: "TimeSpan".to_string(),
                        fields: {
                            let mut f = std::collections::HashMap::new();
                            f.insert("days".to_string(), Value::Integer(days));
                            f.insert("hours".to_string(), Value::Integer(hours));
                            f.insert("minutes".to_string(), Value::Integer(minutes));
                            f.insert("seconds".to_string(), Value::Integer(seconds));
                            f.insert("milliseconds".to_string(), Value::Integer(ms));
                            f.insert("totaldays".to_string(), Value::Double(total_seconds / 86400.0));
                            f.insert("totalhours".to_string(), Value::Double(total_seconds / 3600.0));
                            f.insert("totalminutes".to_string(), Value::Double(total_seconds / 60.0));
                            f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                            f.insert("totalmilliseconds".to_string(), Value::Double(total_seconds * 1000.0));
                            f
                        },
                    };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
                }

                // New Uri(string)
                if class_name == "uri" || class_name == "system.uri" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let url = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Uri".to_string()));
                    fields.insert("absoluteuri".to_string(), Value::String(url.clone()));
                    fields.insert("originalstring".to_string(), Value::String(url.clone()));
                    // Parse components
                    if let Some(scheme_end) = url.find("://") {
                        fields.insert("scheme".to_string(), Value::String(url[..scheme_end].to_string()));
                        let rest = &url[scheme_end + 3..];
                        let (host_port, path) = rest.split_once('/').unwrap_or((rest, ""));
                        let (host, port_str) = host_port.split_once(':').unwrap_or((host_port, ""));
                        fields.insert("host".to_string(), Value::String(host.to_string()));
                        fields.insert("port".to_string(), Value::Integer(port_str.parse().unwrap_or(if url.starts_with("https") { 443 } else { 80 })));
                        fields.insert("absolutepath".to_string(), Value::String(format!("/{}", path)));
                        fields.insert("pathandquery".to_string(), Value::String(format!("/{}", path)));
                    } else {
                        fields.insert("scheme".to_string(), Value::String(String::new()));
                        fields.insert("host".to_string(), Value::String(String::new()));
                        fields.insert("port".to_string(), Value::Integer(0));
                        fields.insert("absolutepath".to_string(), Value::String(url.clone()));
                        fields.insert("pathandquery".to_string(), Value::String(url.clone()));
                    }
                    let obj = crate::value::ObjectData { class_name: "Uri".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // New FileInfo(path)
                if class_name == "fileinfo" || class_name == "system.io.fileinfo" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let path = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                    let p = std::path::Path::new(&path);
                    let meta = std::fs::metadata(&path);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("FileInfo".to_string()));
                    fields.insert("fullname".to_string(), Value::String(std::fs::canonicalize(&path).unwrap_or_else(|_| p.to_path_buf()).to_string_lossy().to_string()));
                    fields.insert("name".to_string(), Value::String(p.file_name().unwrap_or_default().to_string_lossy().to_string()));
                    fields.insert("extension".to_string(), Value::String(p.extension().map(|e| format!(".{}", e.to_string_lossy())).unwrap_or_default()));
                    fields.insert("directoryname".to_string(), Value::String(p.parent().unwrap_or(std::path::Path::new("")).to_string_lossy().to_string()));
                    fields.insert("exists".to_string(), Value::Boolean(p.exists() && p.is_file()));
                    if let Ok(m) = &meta {
                        fields.insert("length".to_string(), Value::Long(m.len() as i64));
                        fields.insert("isreadonly".to_string(), Value::Boolean(m.permissions().readonly()));
                    } else {
                        fields.insert("length".to_string(), Value::Long(0));
                        fields.insert("isreadonly".to_string(), Value::Boolean(false));
                    }
                    let obj = crate::value::ObjectData { class_name: "FileInfo".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // New DirectoryInfo(path)
                if class_name == "directoryinfo" || class_name == "system.io.directoryinfo" {
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;
                    let path = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                    let p = std::path::Path::new(&path);
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("DirectoryInfo".to_string()));
                    fields.insert("fullname".to_string(), Value::String(std::fs::canonicalize(&path).unwrap_or_else(|_| p.to_path_buf()).to_string_lossy().to_string()));
                    fields.insert("name".to_string(), Value::String(p.file_name().unwrap_or_default().to_string_lossy().to_string()));
                    fields.insert("exists".to_string(), Value::Boolean(p.is_dir()));
                    fields.insert("parent".to_string(), Value::String(p.parent().unwrap_or(std::path::Path::new("")).to_string_lossy().to_string()));
                    let obj = crate::value::ObjectData { class_name: "DirectoryInfo".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // System.Random
                if class_name == "random" || class_name == "system.random" {
                    let seed = if !ctor_args.is_empty() {
                        self.evaluate_expr(&ctor_args[0])?.as_integer()? as u64
                    } else {
                        std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_nanos() as u64
                    };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Random".to_string()));
                    fields.insert("__seed".to_string(), Value::Long(seed as i64));
                    fields.insert("__counter".to_string(), Value::Long(0));
                    let obj = crate::value::ObjectData { class_name: "Random".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // System.Diagnostics.Stopwatch
                if class_name == "stopwatch" || class_name == "system.diagnostics.stopwatch" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Stopwatch".to_string()));
                    fields.insert("isrunning".to_string(), Value::Boolean(false));
                    fields.insert("elapsedmilliseconds".to_string(), Value::Long(0));
                    fields.insert("__start_ms".to_string(), Value::Long(0));
                    fields.insert("__accumulated_ms".to_string(), Value::Long(0));
                    let obj = crate::value::ObjectData { class_name: "Stopwatch".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // System.ComponentModel.BackgroundWorker (standalone, not WinForms)
                if class_name == "backgroundworker" || class_name == "system.componentmodel.backgroundworker" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("BackgroundWorker".to_string()));
                    fields.insert("isbusy".to_string(), Value::Boolean(false));
                    fields.insert("cancellationpending".to_string(), Value::Boolean(false));
                    fields.insert("workerreportsprogress".to_string(), Value::Boolean(false));
                    fields.insert("workersupportscancellation".to_string(), Value::Boolean(false));
                    // Event handler name storage (populated by AddHandler or direct assignment)
                    fields.insert("__dowork_handler".to_string(), Value::Nothing);
                    fields.insert("__progresschanged_handler".to_string(), Value::Nothing);
                    fields.insert("__runworkercompleted_handler".to_string(), Value::Nothing);
                    let obj = crate::value::ObjectData { class_name: "BackgroundWorker".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // System.IO.StreamReader
                if class_name == "streamreader" || class_name == "system.io.streamreader" {
                    if ctor_args.is_empty() {
                        return Err(RuntimeError::Custom("StreamReader requires a file path or stream".to_string()));
                    }
                    let arg0 = self.evaluate_expr(&ctor_args[0])?;
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("StreamReader".to_string()));
                    fields.insert("__closed".to_string(), Value::Boolean(false));
                    // Check if arg is a NetworkStream (has __socket_id)
                    if let Value::Object(obj_ref) = &arg0 {
                        let b = obj_ref.borrow();
                        if let Some(Value::Long(sid)) = b.fields.get("__socket_id") {
                            if *sid > 0 {
                                // Network-backed StreamReader
                                fields.insert("__socket_id".to_string(), Value::Long(*sid));
                                fields.insert("__content".to_string(), Value::String(String::new()));
                                fields.insert("__position".to_string(), Value::Integer(0));
                                let obj = crate::value::ObjectData { class_name: "StreamReader".to_string(), fields };
                                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                            }
                        }
                    }
                    // File-backed StreamReader
                    let path = arg0.as_string();
                    let content = std::fs::read_to_string(&path)
                        .map_err(|e| RuntimeError::Custom(format!("StreamReader: {}", e)))?;
                    fields.insert("__content".to_string(), Value::String(content));
                    fields.insert("__position".to_string(), Value::Integer(0));
                    let obj = crate::value::ObjectData { class_name: "StreamReader".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // System.IO.StreamWriter
                if class_name == "streamwriter" || class_name == "system.io.streamwriter" {
                    if ctor_args.is_empty() {
                        return Err(RuntimeError::Custom("StreamWriter requires a file path or stream".to_string()));
                    }
                    let arg0 = self.evaluate_expr(&ctor_args[0])?;
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("StreamWriter".to_string()));
                    fields.insert("__closed".to_string(), Value::Boolean(false));
                    // Check if arg is a NetworkStream (has __socket_id)
                    if let Value::Object(obj_ref) = &arg0 {
                        let b = obj_ref.borrow();
                        if let Some(Value::Long(sid)) = b.fields.get("__socket_id") {
                            if *sid > 0 {
                                // Network-backed StreamWriter
                                fields.insert("__socket_id".to_string(), Value::Long(*sid));
                                fields.insert("__buffer".to_string(), Value::String(String::new()));
                                fields.insert("autoflush".to_string(), Value::Boolean(false));
                                let obj = crate::value::ObjectData { class_name: "StreamWriter".to_string(), fields };
                                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                            }
                        }
                    }
                    // File-backed StreamWriter
                    let path = arg0.as_string();
                    let append = if ctor_args.len() >= 2 {
                        self.evaluate_expr(&ctor_args[1])?.as_bool()?
                    } else {
                        false
                    };
                    fields.insert("__path".to_string(), Value::String(path));
                    fields.insert("__buffer".to_string(), Value::String(String::new()));
                    fields.insert("__append".to_string(), Value::Boolean(append));
                    let obj = crate::value::ObjectData { class_name: "StreamWriter".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // Regex instance: New Regex(pattern) or New Regex(pattern, options)
                if class_name == "regex" || class_name == "system.text.regularexpressions.regex" {
                    let pattern = if !ctor_args.is_empty() {
                        self.evaluate_expr(&ctor_args[0])?.as_string()
                    } else {
                        return Err(RuntimeError::Custom("Regex requires a pattern".to_string()));
                    };
                    let options = if ctor_args.len() >= 2 {
                        self.evaluate_expr(&ctor_args[1])?.as_integer()? // RegexOptions flags
                    } else {
                        0
                    };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("Regex".to_string()));
                    fields.insert("__pattern".to_string(), Value::String(pattern));
                    fields.insert("__options".to_string(), Value::Integer(options));
                    let obj = crate::value::ObjectData { class_name: "Regex".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // StringContent (System.Net.Http.StringContent) — for PostAsync
                if class_name == "stringcontent" || class_name == "system.net.http.stringcontent" {
                    let content = if !ctor_args.is_empty() {
                        self.evaluate_expr(&ctor_args[0])?.as_string()
                    } else {
                        String::new()
                    };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("StringContent".to_string()));
                    fields.insert("body".to_string(), Value::String(content));
                    if ctor_args.len() >= 3 {
                        let media_type = self.evaluate_expr(&ctor_args[2])?.as_string();
                        fields.insert("mediatype".to_string(), Value::String(media_type));
                    }
                    let obj = crate::value::ObjectData {
                        class_name: "StringContent".to_string(),
                        fields,
                    };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ===== DATA ACCESS CONSTRUCTORS =====
                // ADODB.Connection / SqlConnection / OleDbConnection
                if class_name == "adodb.connection" || class_name == "connection"
                    || class_name == "sqlconnection" || class_name == "system.data.sqlclient.sqlconnection"
                    || class_name == "oledbconnection" || class_name == "system.data.oledb.oledbconnection"
                    || class_name == "mysqlconnection" || class_name == "mysql.data.mysqlclient.mysqlconnection"
                    || class_name == "npgsqlconnection" || class_name == "npgsql.npgsqlconnection"
                {
                    let conn_str = if !ctor_args.is_empty() {
                        self.evaluate_expr(&ctor_args[0])?.as_string()
                    } else {
                        String::new()
                    };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("DbConnection".to_string()));
                    fields.insert("__conn_id".to_string(), Value::Long(0)); // 0 = not yet opened
                    fields.insert("connectionstring".to_string(), Value::String(conn_str.clone()));
                    fields.insert("state".to_string(), Value::Integer(0)); // 0=Closed, 1=Open
                    let obj = crate::value::ObjectData { class_name: "DbConnection".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ADODB.Command / SqlCommand / OleDbCommand
                if class_name == "adodb.command" || class_name == "command"
                    || class_name == "sqlcommand" || class_name == "system.data.sqlclient.sqlcommand"
                    || class_name == "oledbcommand" || class_name == "system.data.oledb.oledbcommand"
                    || class_name == "mysqlcommand" || class_name == "mysql.data.mysqlclient.mysqlcommand"
                    || class_name == "npgsqlcommand" || class_name == "npgsql.npgsqlcommand"
                {
                    let sql = if !ctor_args.is_empty() {
                        self.evaluate_expr(&ctor_args[0])?.as_string()
                    } else {
                        String::new()
                    };
                    let conn_id = if ctor_args.len() >= 2 {
                        // Second arg is a connection object — extract its __conn_id
                        let conn_val = self.evaluate_expr(&ctor_args[1])?;
                        if let Value::Object(cr) = &conn_val {
                            cr.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l) } else { None })
                                .unwrap_or(0)
                        } else { 0 }
                    } else { 0 };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("DbCommand".to_string()));
                    fields.insert("__conn_id".to_string(), Value::Long(conn_id));
                    fields.insert("commandtext".to_string(), Value::String(sql));
                    fields.insert("commandtype".to_string(), Value::Integer(1)); // 1=Text
                    // Parameters collection: stores Vec<(name, value)> as a Collection of pairs
                    fields.insert("__parameters".to_string(), Value::Collection(
                        std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))
                    ));
                    let obj = crate::value::ObjectData { class_name: "DbCommand".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // DataAdapter (ADO.NET) — SqlDataAdapter / OleDbDataAdapter
                if class_name == "sqldataadapter" || class_name == "system.data.sqlclient.sqldataadapter"
                    || class_name == "oledbdataadapter" || class_name == "system.data.oledb.oledbdataadapter"
                    || class_name == "mysqldataadapter" || class_name == "dataadapter"
                {
                    let sql = if !ctor_args.is_empty() {
                        self.evaluate_expr(&ctor_args[0])?.as_string()
                    } else { String::new() };
                    let conn_id = if ctor_args.len() >= 2 {
                        let conn_val = self.evaluate_expr(&ctor_args[1])?;
                        if let Value::Object(cr) = &conn_val {
                            cr.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l) } else { None })
                                .unwrap_or(0)
                        } else { 0 }
                    } else { 0 };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("DataAdapter".to_string()));
                    fields.insert("__conn_id".to_string(), Value::Long(conn_id));
                    fields.insert("selectcommandtext".to_string(), Value::String(sql));
                    let obj = crate::value::ObjectData { class_name: "DataAdapter".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ADODB.Recordset
                if class_name == "adodb.recordset" || class_name == "recordset" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("DbRecordset".to_string()));
                    fields.insert("__rs_id".to_string(), Value::Long(0)); // Not yet populated
                    fields.insert("__conn_id".to_string(), Value::Long(0));
                    fields.insert("source".to_string(), Value::String(String::new()));
                    let obj = crate::value::ObjectData { class_name: "DbRecordset".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // DataTable (ADO.NET)
                if class_name == "datatable" || class_name == "system.data.datatable" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("DataTable".to_string()));
                    fields.insert("__rs_id".to_string(), Value::Long(0));
                    fields.insert("tablename".to_string(), Value::String(
                        if !ctor_args.is_empty() { self.evaluate_expr(&ctor_args[0])?.as_string() } else { String::new() }
                    ));
                    let obj = crate::value::ObjectData { class_name: "DataTable".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // DataSet (ADO.NET)
                if class_name == "dataset" || class_name == "system.data.dataset" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("DataSet".to_string()));
                    fields.insert("datasetname".to_string(), Value::String(
                        if !ctor_args.is_empty() { self.evaluate_expr(&ctor_args[0])?.as_string() } else { "NewDataSet".to_string() }
                    ));
                    // Tables collection: Vec<DataTable> stored as Array
                    fields.insert("__tables".to_string(), Value::Array(Vec::new()));
                    let obj = crate::value::ObjectData { class_name: "DataSet".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // BindingSource (System.Windows.Forms)
                if class_name == "bindingsource" || class_name == "system.windows.forms.bindingsource" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("BindingSource".to_string()));
                    fields.insert("__datasource".to_string(), Value::Nothing);
                    fields.insert("datamember".to_string(), Value::String(String::new()));
                    fields.insert("position".to_string(), Value::Integer(0));
                    fields.insert("__bindings".to_string(), Value::Array(Vec::new())); // tracks bound controls
                    fields.insert("name".to_string(), Value::String(String::new()));
                    fields.insert("filter".to_string(), Value::String(String::new()));
                    fields.insert("sort".to_string(), Value::String(String::new()));
                    let obj = crate::value::ObjectData { class_name: "BindingSource".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // ADODB.Parameter
                if class_name == "adodb.parameter" || class_name == "parameter" {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("DbParameter".to_string()));
                    fields.insert("name".to_string(), Value::String(
                        if !ctor_args.is_empty() { self.evaluate_expr(&ctor_args[0])?.as_string() } else { String::new() }
                    ));
                    fields.insert("value".to_string(), Value::Nothing);
                    fields.insert("direction".to_string(), Value::Integer(1)); // adParamInput
                    fields.insert("type".to_string(), Value::Integer(200)); // adVarChar
                    fields.insert("size".to_string(), Value::Integer(0));
                    let obj = crate::value::ObjectData { class_name: "DbParameter".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                let resolved_key = self.resolve_class_key(&class_name);
                if let Some(class_decl) = resolved_key.as_ref().and_then(|k| self.classes.get(k).cloned()) {
                    // Enforce MustInherit — cannot instantiate abstract classes
                    if class_decl.is_must_inherit {
                        return Err(RuntimeError::Custom(format!(
                            "Cannot create an instance of MustInherit class '{}'",
                            class_decl.name.as_str()
                        )));
                    }

                    // Collect fields from hierarchy
                    let fields = self.collect_fields(&class_name);

                    let obj_data = crate::value::ObjectData {
                        class_name: class_decl.name.as_str().to_string(),
                        fields,
                    };

                    let obj_ref = std::rc::Rc::new(std::cell::RefCell::new(obj_data));

                    // Evaluate constructor arguments
                    let arg_values: Result<Vec<_>, _> = ctor_args.iter().map(|e| self.evaluate_expr(e)).collect();
                    let arg_values = arg_values?;

                    // Call Sub New if it exists in hierarchy
                    let new_method = self.find_method(&class_name, "new");

                    if let Some(method) = new_method {
                        // Check if the derived constructor body contains a
                        // MyBase.New() call.  If not, auto-call the base class
                        // Sub New first (VB.NET inserts this automatically).
                        let method_body = match &method {
                            vybe_parser::ast::decl::MethodDecl::Sub(s) => &s.body,
                            vybe_parser::ast::decl::MethodDecl::Function(f) => &f.body,
                        };
                        let has_mybase_new = body_contains_mybase_new(method_body);
                        if !has_mybase_new {
                            if let Some(base_new) = self.find_method_in_base(&class_name, "new") {
                                match base_new {
                                    vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                                        let _ = self.call_user_sub(&s, &[], Some(obj_ref.clone()));
                                    }
                                    vybe_parser::ast::decl::MethodDecl::Function(f) => {
                                        let _ = self.call_user_function(&f, &[], Some(obj_ref.clone()));
                                    }
                                }
                            }
                        }

                        match method {
                            vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                                self.call_user_sub(&s, &arg_values, Some(obj_ref.clone()))?;
                            }
                            vybe_parser::ast::decl::MethodDecl::Function(f) => {
                                self.call_user_function(&f, &arg_values, Some(obj_ref.clone()))?;
                            }
                        }
                    } else {
                        // No Sub New found: for WinForms classes (inheriting System.Windows.Forms.Form),
                        // auto-call InitializeComponent if it exists
                        let inherits_form = class_decl.inherits.as_ref().map_or(false, |t| {
                            match t {
                                vybe_parser::VBType::Custom(n) => n.to_lowercase().contains("form"),
                                _ => false,
                            }
                        });
                        if inherits_form {
                            if let Some(init_method) = self.find_method(&class_name, "InitializeComponent") {
                                match init_method {
                                    vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                                        if let Err(e) = self.call_user_sub(&s, &[], Some(obj_ref.clone())) {
                                            eprintln!("[InitializeComponent error] {}", e);
                                        }
                                    }
                                    vybe_parser::ast::decl::MethodDecl::Function(f) => {
                                        let _ = self.call_user_function(&f, &[], Some(obj_ref.clone()));
                                    }
                                }
                            }
                        }
                    }

                    Ok(Value::Object(obj_ref))
                } else {
                    // If the class is unknown, return Nothing instead of error to keep VB code running
                    return Ok(Value::Nothing);
                }
            }

            Expression::NewFromInitializer(class_id, ctor_args, init_elements) => {
                // New List(Of T) From { expr, expr, ... }
                // Create the collection then add each element
                let obj = self.evaluate_expr(&Expression::New(class_id.clone(), ctor_args.clone()))?;
                for elem_expr in init_elements {
                    let elem_val = self.evaluate_expr(elem_expr)?;
                    // Add to the collection
                    match &obj {
                        Value::Collection(al) => { al.borrow_mut().add(elem_val); }
                        Value::Object(obj_ref) => {
                            // For List, Dictionary etc. — push into __items
                            let mut b = obj_ref.borrow_mut();
                            let items = b.fields.entry("__items".to_string())
                                .or_insert_with(|| Value::Array(Vec::new()));
                            if let Value::Array(arr) = items {
                                arr.push(elem_val);
                            }
                            // Update count
                            if let Some(Value::Array(arr)) = b.fields.get("__items") {
                                let count = arr.len() as i32;
                                b.fields.insert("count".to_string(), Value::Integer(count));
                            }
                        }
                        _ => {}
                    }
                }
                Ok(obj)
            }

            Expression::NewWithInitializer(class_id, ctor_args, members) => {
                // New Type() With { .Prop = expr, ... }
                // Create the object then set each property
                let obj = self.evaluate_expr(&Expression::New(class_id.clone(), ctor_args.clone()))?;
                if let Value::Object(obj_ref) = &obj {
                    for (prop_name, prop_expr) in members {
                        let val = self.evaluate_expr(prop_expr)?;
                        obj_ref.borrow_mut().fields.insert(prop_name.to_lowercase(), val);
                    }
                }
                Ok(obj)
            }
            
            Expression::MemberAccess(obj, member) => {
                // Handle known WinForms enum/namespace values from designer code
                let full_path = format!("{}.{}", self.expr_to_string(obj), member.as_str()).to_lowercase();
                if full_path.contains("system.windows.forms.") 
                    || full_path.contains("system.drawing.")
                    || full_path.contains("system.componentmodel.") {
                    return Ok(Value::Nothing);
                }

                // Try static/qualified property access (e.g., Environment.CurrentDirectory, Math.PI)
                match full_path.as_str() {
                    "environment.currentdirectory" => return Ok(Value::String(std::env::current_dir().unwrap_or_default().to_string_lossy().to_string())),
                    "environment.machinename" => {
                        let name = std::env::var("HOSTNAME")
                            .or_else(|_| std::env::var("COMPUTERNAME"))
                            .unwrap_or_else(|_| {
                                std::process::Command::new("hostname").output()
                                    .map(|o| String::from_utf8_lossy(&o.stdout).trim().to_string())
                                    .unwrap_or_else(|_| "localhost".to_string())
                            });
                        return Ok(Value::String(name));
                    }
                    "environment.username" => return Ok(Value::String(std::env::var("USER").or_else(|_| std::env::var("USERNAME")).unwrap_or_default())),
                    "environment.osversion" => {
                        #[cfg(target_os = "macos")]
                        return Ok(Value::String("Mac OS X".to_string()));
                        #[cfg(target_os = "windows")]
                        return Ok(Value::String("Microsoft Windows NT".to_string()));
                        #[cfg(target_os = "linux")]
                        return Ok(Value::String("Unix".to_string()));
                        #[cfg(not(any(target_os = "macos", target_os = "windows", target_os = "linux")))]
                        return Ok(Value::String("Unknown".to_string()));
                    }
                    "environment.processorcount" => return Ok(Value::Integer(std::thread::available_parallelism().map(|n| n.get()).unwrap_or(1) as i32)),
                    "environment.is64bitoperatingsystem" => return Ok(Value::Boolean(cfg!(target_pointer_width = "64"))),
                    "environment.newline" => return Ok(Value::String("\n".to_string())),
                    "environment.tickcount" | "environment.tickcount64" => {
                        let ms = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_millis() as i64;
                        return Ok(Value::Long(ms));
                    }
                    "environment.version" => return Ok(Value::String("4.0.0".to_string())),
                    "guid.empty" => return Ok(Value::String("00000000-0000-0000-0000-000000000000".to_string())),
                    "path.directoryseparatorchar" => return Ok(Value::String("/".to_string())),
                    // ADO.NET CommandType enum
                    "commandtype.text" => return Ok(Value::Integer(1)),
                    "commandtype.storedprocedure" => return Ok(Value::Integer(4)),
                    "commandtype.tabledirect" => return Ok(Value::Integer(512)),
                    // ADO.NET ConnectionState enum
                    "connectionstate.open" => return Ok(Value::Integer(1)),
                    "connectionstate.closed" => return Ok(Value::Integer(0)),
                    "connectionstate.connecting" => return Ok(Value::Integer(2)),
                    "connectionstate.broken" => return Ok(Value::Integer(16)),
                    // ADODB constants
                    "adcmdtext" => return Ok(Value::Integer(1)),
                    "adcmdstoredproc" => return Ok(Value::Integer(4)),
                    "adcmdtable" => return Ok(Value::Integer(2)),
                    // ADODB Parameter Direction
                    "adparaminput" => return Ok(Value::Integer(1)),
                    "adparamoutput" => return Ok(Value::Integer(2)),
                    "adparaminputoutput" => return Ok(Value::Integer(3)),
                    "adparamreturnvalue" => return Ok(Value::Integer(4)),
                    // ADODB Data Types
                    "advarchar" | "adodb.datatypeenum.advarchar" => return Ok(Value::Integer(200)),
                    "adinteger" | "adodb.datatypeenum.adinteger" => return Ok(Value::Integer(3)),
                    "adboolean" | "adodb.datatypeenum.adboolean" => return Ok(Value::Integer(11)),
                    "addouble" | "adodb.datatypeenum.addouble" => return Ok(Value::Integer(5)),
                    "addate" | "adodb.datatypeenum.addate" => return Ok(Value::Integer(7)),
                    // DBNull.Value
                    "dbnull.value" | "system.dbnull.value" => return Ok(Value::Nothing),
                    // DateTime static properties
                    "datetime.now" | "system.datetime.now" => return Ok(Value::Date(now_ole())),
                    "datetime.today" | "system.datetime.today" => return Ok(Value::Date(today_ole())),
                    "datetime.utcnow" | "system.datetime.utcnow" => return Ok(Value::Date(utcnow_ole())),
                    "datetime.minvalue" | "system.datetime.minvalue" => {
                        return Ok(Value::Date(ymd_to_ole(1, 1, 1, 0, 0, 0)));
                    }
                    "datetime.maxvalue" | "system.datetime.maxvalue" => {
                        return Ok(Value::Date(ymd_to_ole(9999, 12, 31, 23, 59, 59)));
                    }
                    // String static properties
                    "string.empty" | "system.string.empty" => return Ok(Value::String(String::new())),
                    // EventArgs.Empty
                    "eventargs.empty" | "system.eventargs.empty" => return Ok(Self::make_event_args()),
                    // Int32/Double limits
                    "integer.maxvalue" | "int32.maxvalue" | "system.int32.maxvalue" => return Ok(Value::Integer(i32::MAX)),
                    "integer.minvalue" | "int32.minvalue" | "system.int32.minvalue" => return Ok(Value::Integer(i32::MIN)),
                    "long.maxvalue" | "int64.maxvalue" | "system.int64.maxvalue" => return Ok(Value::Long(i64::MAX)),
                    "long.minvalue" | "int64.minvalue" | "system.int64.minvalue" => return Ok(Value::Long(i64::MIN)),
                    "double.maxvalue" | "system.double.maxvalue" => return Ok(Value::Double(f64::MAX)),
                    "double.minvalue" | "system.double.minvalue" => return Ok(Value::Double(f64::MIN)),
                    "double.nan" | "system.double.nan" => return Ok(Value::Double(f64::NAN)),
                    "double.positiveinfinity" | "system.double.positiveinfinity" => return Ok(Value::Double(f64::INFINITY)),
                    "double.negativeinfinity" | "system.double.negativeinfinity" => return Ok(Value::Double(f64::NEG_INFINITY)),
                    "single.maxvalue" | "system.single.maxvalue" => return Ok(Value::Single(f32::MAX)),
                    "single.minvalue" | "system.single.minvalue" => return Ok(Value::Single(f32::MIN)),
                    // TimeSpan.Zero
                    "timespan.zero" | "system.timespan.zero" => {
                        let obj_data = crate::value::ObjectData {
                            class_name: "TimeSpan".to_string(),
                            fields: {
                                let mut f = std::collections::HashMap::new();
                                f.insert("days".to_string(), Value::Integer(0));
                                f.insert("hours".to_string(), Value::Integer(0));
                                f.insert("minutes".to_string(), Value::Integer(0));
                                f.insert("seconds".to_string(), Value::Integer(0));
                                f.insert("milliseconds".to_string(), Value::Integer(0));
                                f.insert("totalseconds".to_string(), Value::Double(0.0));
                                f.insert("totalmilliseconds".to_string(), Value::Double(0.0));
                                f
                            },
                        };
                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
                    }

                    // ===== WinForms Enum Constants =====
                    // MouseButtons
                    "mousebuttons.left" | "system.windows.forms.mousebuttons.left" => return Ok(Value::Integer(0x100000)),
                    "mousebuttons.right" | "system.windows.forms.mousebuttons.right" => return Ok(Value::Integer(0x200000)),
                    "mousebuttons.middle" | "system.windows.forms.mousebuttons.middle" => return Ok(Value::Integer(0x400000)),
                    "mousebuttons.none" | "system.windows.forms.mousebuttons.none" => return Ok(Value::Integer(0)),
                    // DialogResult
                    "dialogresult.ok" | "system.windows.forms.dialogresult.ok" | "windows.forms.dialogresult.ok" => return Ok(Value::Integer(1)),
                    "dialogresult.cancel" | "system.windows.forms.dialogresult.cancel" | "windows.forms.dialogresult.cancel" => return Ok(Value::Integer(2)),
                    "dialogresult.abort" | "system.windows.forms.dialogresult.abort" => return Ok(Value::Integer(3)),
                    "dialogresult.retry" | "system.windows.forms.dialogresult.retry" => return Ok(Value::Integer(4)),
                    "dialogresult.ignore" | "system.windows.forms.dialogresult.ignore" => return Ok(Value::Integer(5)),
                    "dialogresult.yes" | "system.windows.forms.dialogresult.yes" | "windows.forms.dialogresult.yes" => return Ok(Value::Integer(6)),
                    "dialogresult.no" | "system.windows.forms.dialogresult.no" | "windows.forms.dialogresult.no" => return Ok(Value::Integer(7)),
                    "dialogresult.none" | "system.windows.forms.dialogresult.none" => return Ok(Value::Integer(0)),
                    // MessageBoxButtons
                    "messageboxbuttons.ok" | "system.windows.forms.messageboxbuttons.ok" => return Ok(Value::Integer(0)),
                    "messageboxbuttons.okcancel" | "system.windows.forms.messageboxbuttons.okcancel" => return Ok(Value::Integer(1)),
                    "messageboxbuttons.abortretryignore" | "system.windows.forms.messageboxbuttons.abortretryignore" => return Ok(Value::Integer(2)),
                    "messageboxbuttons.yesnocancel" | "system.windows.forms.messageboxbuttons.yesnocancel" => return Ok(Value::Integer(3)),
                    "messageboxbuttons.yesno" | "system.windows.forms.messageboxbuttons.yesno" => return Ok(Value::Integer(4)),
                    "messageboxbuttons.retrycancel" | "system.windows.forms.messageboxbuttons.retrycancel" => return Ok(Value::Integer(5)),
                    // MessageBoxIcon
                    "messageboxicon.none" | "system.windows.forms.messageboxicon.none" => return Ok(Value::Integer(0)),
                    "messageboxicon.error" | "system.windows.forms.messageboxicon.error" => return Ok(Value::Integer(16)),
                    "messageboxicon.warning" | "system.windows.forms.messageboxicon.warning" => return Ok(Value::Integer(48)),
                    "messageboxicon.information" | "system.windows.forms.messageboxicon.information" => return Ok(Value::Integer(64)),
                    "messageboxicon.question" | "system.windows.forms.messageboxicon.question" => return Ok(Value::Integer(32)),
                    // DockStyle
                    "dockstyle.none" | "system.windows.forms.dockstyle.none" => return Ok(Value::Integer(0)),
                    "dockstyle.top" | "system.windows.forms.dockstyle.top" => return Ok(Value::Integer(1)),
                    "dockstyle.bottom" | "system.windows.forms.dockstyle.bottom" => return Ok(Value::Integer(2)),
                    "dockstyle.left" | "system.windows.forms.dockstyle.left" => return Ok(Value::Integer(3)),
                    "dockstyle.right" | "system.windows.forms.dockstyle.right" => return Ok(Value::Integer(4)),
                    "dockstyle.fill" | "system.windows.forms.dockstyle.fill" => return Ok(Value::Integer(5)),
                    // AnchorStyles
                    "anchorstyles.none" | "system.windows.forms.anchorstyles.none" => return Ok(Value::Integer(0)),
                    "anchorstyles.top" | "system.windows.forms.anchorstyles.top" => return Ok(Value::Integer(1)),
                    "anchorstyles.bottom" | "system.windows.forms.anchorstyles.bottom" => return Ok(Value::Integer(2)),
                    "anchorstyles.left" | "system.windows.forms.anchorstyles.left" => return Ok(Value::Integer(4)),
                    "anchorstyles.right" | "system.windows.forms.anchorstyles.right" => return Ok(Value::Integer(8)),
                    // FormBorderStyle
                    "formborderstyle.none" | "system.windows.forms.formborderstyle.none" => return Ok(Value::Integer(0)),
                    "formborderstyle.fixedsingle" | "system.windows.forms.formborderstyle.fixedsingle" => return Ok(Value::Integer(1)),
                    "formborderstyle.fixed3d" | "system.windows.forms.formborderstyle.fixed3d" => return Ok(Value::Integer(2)),
                    "formborderstyle.fixeddialog" | "system.windows.forms.formborderstyle.fixeddialog" => return Ok(Value::Integer(3)),
                    "formborderstyle.sizable" | "system.windows.forms.formborderstyle.sizable" => return Ok(Value::Integer(4)),
                    "formborderstyle.fixedtoolwindow" | "system.windows.forms.formborderstyle.fixedtoolwindow" => return Ok(Value::Integer(5)),
                    "formborderstyle.sizabletoolwindow" | "system.windows.forms.formborderstyle.sizabletoolwindow" => return Ok(Value::Integer(6)),
                    // FormStartPosition
                    "formstartposition.manual" | "system.windows.forms.formstartposition.manual" => return Ok(Value::Integer(0)),
                    "formstartposition.centerscreen" | "system.windows.forms.formstartposition.centerscreen" => return Ok(Value::Integer(1)),
                    "formstartposition.windowsdefaultlocation" | "system.windows.forms.formstartposition.windowsdefaultlocation" => return Ok(Value::Integer(2)),
                    "formstartposition.windowsdefaultbounds" | "system.windows.forms.formstartposition.windowsdefaultbounds" => return Ok(Value::Integer(3)),
                    "formstartposition.centerparent" | "system.windows.forms.formstartposition.centerparent" => return Ok(Value::Integer(4)),
                    // FormWindowState
                    "formwindowstate.normal" | "system.windows.forms.formwindowstate.normal" => return Ok(Value::Integer(0)),
                    "formwindowstate.minimized" | "system.windows.forms.formwindowstate.minimized" => return Ok(Value::Integer(1)),
                    "formwindowstate.maximized" | "system.windows.forms.formwindowstate.maximized" => return Ok(Value::Integer(2)),
                    // Keys enum (common keys)
                    "keys.none" | "system.windows.forms.keys.none" => return Ok(Value::Integer(0)),
                    "keys.enter" | "system.windows.forms.keys.enter" | "keys.return" | "system.windows.forms.keys.return" => return Ok(Value::Integer(13)),
                    "keys.escape" | "system.windows.forms.keys.escape" => return Ok(Value::Integer(27)),
                    "keys.space" | "system.windows.forms.keys.space" => return Ok(Value::Integer(32)),
                    "keys.back" | "system.windows.forms.keys.back" => return Ok(Value::Integer(8)),
                    "keys.tab" | "system.windows.forms.keys.tab" => return Ok(Value::Integer(9)),
                    "keys.delete" | "system.windows.forms.keys.delete" => return Ok(Value::Integer(46)),
                    "keys.insert" | "system.windows.forms.keys.insert" => return Ok(Value::Integer(45)),
                    "keys.up" | "system.windows.forms.keys.up" => return Ok(Value::Integer(38)),
                    "keys.down" | "system.windows.forms.keys.down" => return Ok(Value::Integer(40)),
                    "keys.left" | "system.windows.forms.keys.left" => return Ok(Value::Integer(37)),
                    "keys.right" | "system.windows.forms.keys.right" => return Ok(Value::Integer(39)),
                    "keys.f1" | "system.windows.forms.keys.f1" => return Ok(Value::Integer(112)),
                    "keys.f2" | "system.windows.forms.keys.f2" => return Ok(Value::Integer(113)),
                    "keys.f3" | "system.windows.forms.keys.f3" => return Ok(Value::Integer(114)),
                    "keys.f4" | "system.windows.forms.keys.f4" => return Ok(Value::Integer(115)),
                    "keys.f5" | "system.windows.forms.keys.f5" => return Ok(Value::Integer(116)),
                    "keys.control" | "system.windows.forms.keys.control" | "keys.controlkey" | "system.windows.forms.keys.controlkey" => return Ok(Value::Integer(17)),
                    "keys.shift" | "system.windows.forms.keys.shift" | "keys.shiftkey" | "system.windows.forms.keys.shiftkey" => return Ok(Value::Integer(16)),
                    "keys.alt" | "system.windows.forms.keys.alt" | "keys.menu" | "system.windows.forms.keys.menu" => return Ok(Value::Integer(18)),
                    // Color constants
                    "color.red" | "system.drawing.color.red" => return Ok(Value::Integer(0xFF0000)),
                    "color.green" | "system.drawing.color.green" => return Ok(Value::Integer(0x008000)),
                    "color.blue" | "system.drawing.color.blue" => return Ok(Value::Integer(0x0000FF)),
                    "color.white" | "system.drawing.color.white" => return Ok(Value::Integer(0xFFFFFF)),
                    "color.black" | "system.drawing.color.black" => return Ok(Value::Integer(0x000000)),
                    "color.yellow" | "system.drawing.color.yellow" => return Ok(Value::Integer(0xFFFF00)),
                    "color.gray" | "system.drawing.color.gray" => return Ok(Value::Integer(0x808080)),
                    "color.transparent" | "system.drawing.color.transparent" => return Ok(Value::Integer(0)),
                    // ContentAlignment
                    "contentalignment.middlecenter" | "system.drawing.contentalignment.middlecenter" => return Ok(Value::Integer(32)),
                    "contentalignment.middleleft" | "system.drawing.contentalignment.middleleft" => return Ok(Value::Integer(16)),
                    "contentalignment.middleright" | "system.drawing.contentalignment.middleright" => return Ok(Value::Integer(64)),
                    "contentalignment.topcenter" | "system.drawing.contentalignment.topcenter" => return Ok(Value::Integer(2)),
                    "contentalignment.topleft" | "system.drawing.contentalignment.topleft" => return Ok(Value::Integer(1)),
                    "contentalignment.topright" | "system.drawing.contentalignment.topright" => return Ok(Value::Integer(4)),
                    "contentalignment.bottomcenter" | "system.drawing.contentalignment.bottomcenter" => return Ok(Value::Integer(512)),
                    "contentalignment.bottomleft" | "system.drawing.contentalignment.bottomleft" => return Ok(Value::Integer(256)),
                    "contentalignment.bottomright" | "system.drawing.contentalignment.bottomright" => return Ok(Value::Integer(1024)),
                    // CloseReason
                    "closereason.none" | "system.windows.forms.closereason.none" => return Ok(Value::Integer(0)),
                    "closereason.windowsshutdown" | "system.windows.forms.closereason.windowsshutdown" => return Ok(Value::Integer(1)),
                    "closereason.userclosing" | "system.windows.forms.closereason.userclosing" => return Ok(Value::Integer(3)),
                    "closereason.applicationexitcall" | "system.windows.forms.closereason.applicationexitcall" => return Ok(Value::Integer(6)),

                    _ => {}
                }

                // Check if the full_path corresponds to a registered constant
                if let Ok(val) = self.env.get(&full_path) {
                    return Ok(val.clone());
                }
                // Also check original-case
                let full_path_orig = format!("{}.{}", self.expr_to_string(obj), member.as_str());
                if let Ok(val) = self.env.get(&full_path_orig) {
                    return Ok(val.clone());
                }

                let obj_val = self.evaluate_expr(obj)?;
                
                // Collection Properties
                if let Value::Collection(col_rc) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    match m.as_str() {
                        "count" => return Ok(Value::Integer(col_rc.borrow().count())),
                        "capacity" => return Ok(Value::Integer(col_rc.borrow().capacity())),
                        "isfixedsize" | "isreadonly" => return Ok(Value::Boolean(false)),
                        "issynchronized" => return Ok(Value::Boolean(false)),
                        _ => {}
                    }
                }

                // Queue Properties
                if let Value::Queue(q) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    if m == "count" { return Ok(Value::Integer(q.borrow().count())); }
                }

                // Stack Properties
                if let Value::Stack(s) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    if m == "count" { return Ok(Value::Integer(s.borrow().count())); }
                }

                // HashSet Properties
                if let Value::HashSet(h) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    if m == "count" { return Ok(Value::Integer(h.borrow().count())); }
                }

                // Dictionary Properties
                if let Value::Dictionary(d) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    match m.as_str() {
                        "count" => return Ok(Value::Integer(d.borrow().count())),
                        "keys" => return Ok(Value::Array(d.borrow().keys())),
                        "values" => return Ok(Value::Array(d.borrow().values())),
                        _ => {}
                    }
                }

                // ConcurrentDictionary Properties
                if let Value::ConcurrentDictionary(d) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    match m.as_str() {
                        "count" => return Ok(Value::Integer(d.count())),
                        "keys" => return Ok(Value::Array(d.keys().into_iter().map(Value::String).collect())),
                        "values" => return Ok(Value::Array(d.values())),
                        _ => {}
                    }
                }

                // ConcurrentQueue Properties
                if let Value::ConcurrentQueue(q) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    if m == "count" { return Ok(Value::Integer(q.count())); }
                }

                // ConcurrentStack Properties
                if let Value::ConcurrentStack(s) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    if m == "count" { return Ok(Value::Integer(s.count())); }
                }

                // Array Properties (Length, Count)
                if let Value::Array(arr) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    if m == "length" || m == "count" {
                        return Ok(Value::Integer(arr.len() as i32));
                    }
                }

                // String Properties (Length, Chars)
                if let Value::String(s) = &obj_val {
                    let m = member.as_str().to_lowercase();
                    if m == "length" {
                        return Ok(Value::Integer(s.len() as i32));
                    }
                }

                // DateTime Properties (Year, Month, Day, etc.) — OLE f64
                if let Value::Date(ole_val) = &obj_val {
                    let ndt = ole_to_dt(*ole_val);
                    use chrono::{Datelike, Timelike, NaiveDate};
                    let m = member.as_str().to_lowercase();
                    match m.as_str() {
                        "year" => return Ok(Value::Integer(ndt.year())),
                        "month" => return Ok(Value::Integer(ndt.month() as i32)),
                        "day" => return Ok(Value::Integer(ndt.day() as i32)),
                        "hour" => return Ok(Value::Integer(ndt.hour() as i32)),
                        "minute" => return Ok(Value::Integer(ndt.minute() as i32)),
                        "second" => return Ok(Value::Integer(ndt.second() as i32)),
                        "millisecond" => return Ok(Value::Integer((ndt.nanosecond() / 1_000_000) as i32)),
                        "dayofweek" => return Ok(Value::Integer(ndt.weekday().num_days_from_sunday() as i32)),
                        "dayofyear" => return Ok(Value::Integer(ndt.ordinal() as i32)),
                        "date" => {
                            let d = ndt.date().and_hms_opt(0, 0, 0).unwrap();
                            return Ok(Value::Date(date_to_ole(d)));
                        }
                        "timeofday" => {
                            let total_seconds = (ndt.hour() as f64) * 3600.0 + (ndt.minute() as f64) * 60.0 + (ndt.second() as f64);
                            let obj_data = crate::value::ObjectData {
                                class_name: "TimeSpan".to_string(),
                                fields: {
                                    let mut f = std::collections::HashMap::new();
                                    f.insert("hours".to_string(), Value::Integer(ndt.hour() as i32));
                                    f.insert("minutes".to_string(), Value::Integer(ndt.minute() as i32));
                                    f.insert("seconds".to_string(), Value::Integer(ndt.second() as i32));
                                    f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                                    f.insert("totalmilliseconds".to_string(), Value::Double(total_seconds * 1000.0));
                                    f
                                },
                            };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
                        }
                        "ticks" => {
                            let epoch = NaiveDate::from_ymd_opt(1, 1, 1).unwrap().and_hms_opt(0, 0, 0).unwrap();
                            let diff = ndt.signed_duration_since(epoch);
                            return Ok(Value::Long(diff.num_milliseconds() * 10_000));
                        }
                        _ => {}
                    }
                }

                // Boolean Properties (via string)
                // Integer/Long/Double properties - fallback to env lookup below

                // XML object property access (XmlDocument, XmlElement, XmlAttribute, etc.)
                if crate::builtins::xml::is_xml_object(&obj_val) {
                    return crate::builtins::xml::xml_property_access(&obj_val, member.as_str());
                }

                if let Value::Object(obj_ref) = &obj_val {
                    let class_name_str;
                    {
                        let obj_data = obj_ref.borrow();
                        class_name_str = obj_data.class_name.clone();

                        // Special: Stopwatch.ElapsedMilliseconds — compute live if running
                        if class_name_str == "Stopwatch" && member.as_str().eq_ignore_ascii_case("ElapsedMilliseconds") {
                            let is_running = obj_data.fields.get("isrunning").map(|v| if let Value::Boolean(b) = v { *b } else { false }).unwrap_or(false);
                            let accumulated = obj_data.fields.get("__accumulated_ms").map(|v| if let Value::Long(l) = v { *l } else { 0 }).unwrap_or(0);
                            if is_running {
                                let start = obj_data.fields.get("__start_ms").map(|v| if let Value::Long(l) = v { *l } else { 0 }).unwrap_or(0);
                                let now_ms = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_millis() as i64;
                                return Ok(Value::Long(accumulated + (now_ms - start)));
                            } else {
                                return Ok(Value::Long(accumulated));
                            }
                        }

                        // Special: Guid.ToString() and Guid member access
                        if class_name_str == "Guid" && member.as_str().eq_ignore_ascii_case("ToString") {
                            if let Some(val) = obj_data.fields.get("__value") {
                                return Ok(val.clone());
                            }
                        }

                        // Special: StreamReader.EndOfStream
                        if class_name_str == "StreamReader" && member.as_str().eq_ignore_ascii_case("EndOfStream") {
                            // Network-backed: check __closed flag (can't peek a socket easily)
                            if obj_data.fields.contains_key("__socket_id") {
                                let closed = obj_data.fields.get("__closed").and_then(|v| v.as_bool().ok()).unwrap_or(false);
                                return Ok(Value::Boolean(closed));
                            }
                            let content_len = obj_data.fields.get("__content").map(|v| v.as_string().len()).unwrap_or(0);
                            let pos = obj_data.fields.get("__position").map(|v| if let Value::Integer(i) = v { *i as usize } else { 0 }).unwrap_or(0);
                            return Ok(Value::Boolean(pos >= content_len));
                        }

                        // ===== DATABASE PROPERTY ACCESS =====
                        let db_type = obj_data.fields.get("__type").and_then(|v| {
                            if let Value::String(s) = v { Some(s.clone()) } else { None }
                        }).unwrap_or_default();

                        // DbRecordset properties (ADODB)
                        if db_type == "DbRecordset" {
                            let m = member.as_str().to_lowercase();
                            let rs_id = obj_data.fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            match m.as_str() {
                                "eof" => {
                                    let dam = crate::data_access::get_global_dam();
                                    let is_eof = dam.lock().unwrap().recordsets.get(&rs_id)
                                        .map(|rs| rs.eof()).unwrap_or(true);
                                    return Ok(Value::Boolean(is_eof));
                                }
                                "bof" => {
                                    let dam = crate::data_access::get_global_dam();
                                    let is_bof = dam.lock().unwrap().recordsets.get(&rs_id)
                                        .map(|rs| rs.bof()).unwrap_or(true);
                                    return Ok(Value::Boolean(is_bof));
                                }
                                "recordcount" => {
                                    let dam = crate::data_access::get_global_dam();
                                    let count = dam.lock().unwrap().recordsets.get(&rs_id)
                                        .map(|rs| rs.record_count()).unwrap_or(0);
                                    return Ok(Value::Integer(count));
                                }
                                _ => {}
                            }
                        }

                        // Task properties
                        if db_type == "Task" {
                            let m = member.as_str().to_lowercase();
                            let handle_id = obj_data.fields.get("__handle").map(|v| v.as_string()).unwrap_or_default();
                            if !handle_id.is_empty() {
                                let shared_obj_opt = get_registry().lock().unwrap().shared_objects.get(&handle_id).cloned();
                                if let Some(shared_obj) = shared_obj_opt {
                                    match m.as_str() {
                                        "result" => {
                                            loop {
                                                {
                                                    let lock = shared_obj.lock().unwrap();
                                                    if let Some(crate::value::SharedValue::Boolean(true)) = lock.fields.get("iscompleted") {
                                                        break;
                                                    }
                                                }
                                                std::thread::sleep(std::time::Duration::from_millis(10));
                                            }
                                            let lock = shared_obj.lock().unwrap();
                                            return Ok(lock.fields.get("result").cloned().unwrap_or(crate::value::SharedValue::Nothing).to_value());
                                        }
                                        "iscompleted" => {
                                            let lock = shared_obj.lock().unwrap();
                                            return Ok(lock.fields.get("iscompleted").cloned().unwrap_or(crate::value::SharedValue::Boolean(false)).to_value());
                                        }
                                        "status" => {
                                            let lock = shared_obj.lock().unwrap();
                                            return Ok(lock.fields.get("status").cloned().unwrap_or(crate::value::SharedValue::String("Unknown".to_string())).to_value());
                                        }
                                        _ => {}
                                    }
                                }
                            }
                        }

                        // Thread properties
                        if class_name_str == "Thread" {
                            let m = member.as_str().to_lowercase();
                            let handle_id = obj_data.fields.get("__handle").map(|v| v.as_string()).unwrap_or_default();
                            if !handle_id.is_empty() {
                                let shared_obj_opt = get_registry().lock().unwrap().shared_objects.get(&handle_id).cloned();
                                if let Some(shared_obj) = shared_obj_opt {
                                    match m.as_str() {
                                        "isalive" => {
                                            let lock = shared_obj.lock().unwrap();
                                            return Ok(lock.fields.get("isalive").cloned().unwrap_or(crate::value::SharedValue::Boolean(false)).to_value());
                                        }
                                        "managedthreadid" => {
                                            return Ok(Value::Integer(handle_id.len() as i32));
                                        }
                                        _ => {}
                                    }
                                }
                            }
                        }

                        // Process properties (basic)
                        if class_name_str == "Process" {
                            let m = member.as_str().to_lowercase();
                            if m == "exitcode" {
                                return Ok(obj_data.fields.get("exitcode").cloned().unwrap_or(Value::Integer(0)));
                            }
                            if m == "hasexited" {
                                return Ok(obj_data.fields.get("hasexited").cloned().unwrap_or(Value::Boolean(false)));
                            }
                        }

                        // DbReader properties (ADO.NET)
                        if db_type == "DbReader" {
                            let m = member.as_str().to_lowercase();
                            let rs_id = obj_data.fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            match m.as_str() {
                                "hasrows" => {
                                    let dam = crate::data_access::get_global_dam();
                                    let has = dam.lock().unwrap().recordsets.get(&rs_id)
                                        .map(|rs| !rs.rows.is_empty()).unwrap_or(false);
                                    return Ok(Value::Boolean(has));
                                }
                                "fieldcount" => {
                                    let dam = crate::data_access::get_global_dam();
                                    let count = dam.lock().unwrap().recordsets.get(&rs_id)
                                        .map(|rs| rs.field_count()).unwrap_or(0);
                                    return Ok(Value::Integer(count));
                                }
                                "isclosed" => {
                                    let dam = crate::data_access::get_global_dam();
                                    let closed = !dam.lock().unwrap().recordsets.contains_key(&rs_id);
                                    return Ok(Value::Boolean(closed));
                                }
                                _ => {}
                            }
                        }

                        // DbConnection properties
                        if db_type == "DbConnection" {
                            let m = member.as_str().to_lowercase();
                            match m.as_str() {
                                "state" => {
                                    let state = obj_data.fields.get("state")
                                        .cloned().unwrap_or(Value::Integer(0));
                                    return Ok(state);
                                }
                                "connectionstring" => {
                                    let cs = obj_data.fields.get("connectionstring")
                                        .cloned().unwrap_or(Value::String(String::new()));
                                    return Ok(cs);
                                }
                                "database" => {
                                    // Return database name from connection string
                                    let cs = obj_data.fields.get("connectionstring")
                                        .map(|v| v.as_string()).unwrap_or_default();
                                    let lower = cs.to_lowercase();
                                    let db = if lower.contains("database=") {
                                        cs.split(';').find(|p| p.trim().to_lowercase().starts_with("database="))
                                            .and_then(|p| p.split('=').nth(1))
                                            .unwrap_or("").trim().to_string()
                                    } else { String::new() };
                                    return Ok(Value::String(db));
                                }
                                "serverversion" => {
                                    // Return backend name as version proxy
                                    let conn_id = obj_data.fields.get("__conn_id")
                                        .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                        .unwrap_or(0);
                                    let dam = crate::data_access::get_global_dam();
                                    let dam_lock = dam.lock().unwrap();
                                    let ver = if let Some((provider, _)) = dam_lock.get_connection_info(conn_id) {
                                        provider
                                    } else { "Unknown".to_string() };
                                    return Ok(Value::String(ver));
                                }
                                "provider" | "datasource" => {
                                    let conn_id = obj_data.fields.get("__conn_id")
                                        .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                        .unwrap_or(0);
                                    let dam = crate::data_access::get_global_dam();
                                    let dam_lock = dam.lock().unwrap();
                                    let name = if let Some((provider, _)) = dam_lock.get_connection_info(conn_id) {
                                        provider
                                    } else { "Unknown".to_string() };
                                    return Ok(Value::String(name));
                                }
                                "connectiontimeout" => {
                                    return Ok(Value::Integer(30));
                                }
                                _ => {}
                            }
                        }

                        // DbCommand properties
                        if db_type == "DbCommand" {
                            let m = member.as_str().to_lowercase();
                            if m == "parameters" {
                                // Return a DbParameters proxy object
                                let obj_rc = if let Value::Object(r) = &obj_val { r.clone() } else { unreachable!() };
                                let mut fields = std::collections::HashMap::new();
                                fields.insert("__type".to_string(), Value::String("DbParameters".to_string()));
                                fields.insert("__parent_cmd".to_string(), Value::Object(obj_rc));
                                let obj = crate::value::ObjectData { class_name: "DbParameters".to_string(), fields };
                                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                            }
                            if m == "connection" {
                                // Return connection object ref if stored, else Nothing
                                return Ok(Value::Nothing);
                            }
                            if m == "commandtimeout" {
                                let timeout = obj_data.fields.get("commandtimeout")
                                    .cloned().unwrap_or(Value::Integer(30));
                                return Ok(timeout);
                            }
                            if m == "commandtype" {
                                let ct = obj_data.fields.get("commandtype")
                                    .cloned().unwrap_or(Value::Integer(1));
                                return Ok(ct);
                            }
                            if m == "commandtext" {
                                let ct = obj_data.fields.get("commandtext")
                                    .cloned().unwrap_or(Value::String(String::new()));
                                return Ok(ct);
                            }
                        }

                        // DataTable properties
                        if db_type == "DataTable" {
                            let m = member.as_str().to_lowercase();
                            let rs_id = obj_data.fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            match m.as_str() {
                                "rows" => {
                                    // Return a DataRowCollection proxy
                                    let dam = crate::data_access::get_global_dam();
                                    let dam_lock = dam.lock().unwrap();
                                    if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                        // Build an array of DataRow objects
                                        let mut row_objects = Vec::new();
                                        for (i, db_row) in rs.rows.iter().enumerate() {
                                            let mut flds = std::collections::HashMap::new();
                                            flds.insert("__type".to_string(), Value::String("DataRow".to_string()));
                                            flds.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                            flds.insert("__row_index".to_string(), Value::Integer(i as i32));
                                            // Copy column values as fields
                                            for (ci, col) in db_row.columns.iter().enumerate() {
                                                let v = db_row.values.get(ci).cloned().unwrap_or_default();
                                                flds.insert(col.to_lowercase(), Value::String(v));
                                            }
                                            let obj = crate::value::ObjectData { class_name: "DataRow".to_string(), fields: flds };
                                            row_objects.push(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                        }
                                        return Ok(Value::Array(row_objects));
                                    }
                                    return Ok(Value::Array(Vec::new()));
                                }
                                "columns" => {
                                    let dam = crate::data_access::get_global_dam();
                                    let dam_lock = dam.lock().unwrap();
                                    if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                        let cols: Vec<Value> = rs.columns.iter()
                                            .map(|c| Value::String(c.clone())).collect();
                                        return Ok(Value::Array(cols));
                                    }
                                    return Ok(Value::Array(Vec::new()));
                                }
                                "tablename" => {
                                    return Ok(obj_data.fields.get("tablename")
                                        .cloned().unwrap_or(Value::String(String::new())));
                                }
                                _ => {}
                            }
                        }

                        // DataSet properties
                        if db_type == "DataSet" {
                            let m = member.as_str().to_lowercase();
                            match m.as_str() {
                                "tables" => {
                                    let tables = obj_data.fields.get("__tables")
                                        .cloned().unwrap_or(Value::Array(Vec::new()));
                                    return Ok(tables);
                                }
                                "datasetname" => {
                                    return Ok(obj_data.fields.get("datasetname")
                                        .cloned().unwrap_or(Value::String("NewDataSet".to_string())));
                                }
                                _ => {}
                            }
                        }

                        // BindingSource properties
                        if db_type == "BindingSource" {
                            let m = member.as_str().to_lowercase();
                            match m.as_str() {
                                "datasource" => {
                                    return Ok(obj_data.fields.get("__datasource")
                                        .cloned().unwrap_or(Value::Nothing));
                                }
                                "datamember" => {
                                    return Ok(obj_data.fields.get("datamember")
                                        .cloned().unwrap_or(Value::String(String::new())));
                                }
                                "position" => {
                                    return Ok(obj_data.fields.get("position")
                                        .cloned().unwrap_or(Value::Integer(0)));
                                }
                                "count" => {
                                    // Get row count from underlying data source
                                    let ds = obj_data.fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                                    drop(obj_data);
                                    let count = self.binding_source_row_count(&ds);
                                    return Ok(Value::Integer(count));
                                }
                                "current" => {
                                    // Return the current DataRow based on position
                                    let ds = obj_data.fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                                    let pos = obj_data.fields.get("position")
                                        .and_then(|v| if let Value::Integer(i) = v { Some(*i) } else { None })
                                        .unwrap_or(0);
                                    drop(obj_data);
                                    let row = self.binding_source_get_row(&ds, pos);
                                    return Ok(row);
                                }
                                "filter" => {
                                    return Ok(obj_data.fields.get("filter")
                                        .cloned().unwrap_or(Value::String(String::new())));
                                }
                                "sort" => {
                                    return Ok(obj_data.fields.get("sort")
                                        .cloned().unwrap_or(Value::String(String::new())));
                                }
                                "name" => {
                                    return Ok(obj_data.fields.get("name")
                                        .cloned().unwrap_or(Value::String(String::new())));
                                }
                                "databindings" => {
                                    // Return a DataBindings proxy
                                    let obj_rc = if let Value::Object(r) = &obj_val { r.clone() } else { unreachable!() };
                                    let mut flds = std::collections::HashMap::new();
                                    flds.insert("__type".to_string(), Value::String("DataBindings".to_string()));
                                    flds.insert("__parent".to_string(), Value::Object(obj_rc));
                                    let proxy = crate::value::ObjectData { class_name: "DataBindings".to_string(), fields: flds };
                                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(proxy))));
                                }
                                _ => {}
                            }
                        }

                        // DbParameter properties
                        if db_type == "DbParameter" {
                            let m = member.as_str().to_lowercase();
                            if let Some(val) = obj_data.fields.get(&m) {
                                return Ok(val.clone());
                            }
                        }

                        // DbParameters collection properties
                        if db_type == "DbParameters" {
                            let m = member.as_str().to_lowercase();
                            if m == "count" {
                                let parent_cmd = obj_data.fields.get("__parent_cmd").cloned();
                                drop(obj_data); // must drop borrow before new borrows
                                if let Some(Value::Object(cmd_ref)) = parent_cmd {
                                    let cmd_borrow = cmd_ref.borrow();
                                    if let Some(Value::Collection(coll_rc)) = cmd_borrow.fields.get("__parameters") {
                                        return Ok(Value::Integer(coll_rc.borrow().items.len() as i32));
                                    }
                                }
                                return Ok(Value::Integer(0));
                            }
                        }
                        
                        // 1. Check if it's a field in the map (case-insensitive)
                        // Skip "controls" — it has a dedicated dynamic accessor below
                        let member_lc = member.as_str().to_lowercase();
                        if member_lc != "controls" {
                            if let Some(val) = obj_data.fields.get(&member_lc) {
                                return Ok(val.clone());
                            }
                        }
                    } // Drop borrow

                    // Handle WinForms infrastructure properties that don't exist as real fields
                    let member_lower = member.as_str().to_lowercase();
                    
                    // Computed properties for control objects
                    {
                        let is_ctrl = obj_ref.borrow().fields.get("__is_control")
                            .map(|v| v.as_bool().unwrap_or(false)).unwrap_or(false);
                        if is_ctrl {
                            match member_lower.as_str() {
                                "location" => {
                                    let b = obj_ref.borrow();
                                    let x = b.fields.get("left").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    let y = b.fields.get("top").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    return Ok(Self::make_point(x, y));
                                }
                                "size" => {
                                    let b = obj_ref.borrow();
                                    let w = b.fields.get("width").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    let h = b.fields.get("height").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    return Ok(Self::make_size(w, h));
                                }
                                "clientsize" => {
                                    let b = obj_ref.borrow();
                                    let w = b.fields.get("width").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    let h = b.fields.get("height").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    return Ok(Self::make_size(w, h));
                                }
                                "bounds" => {
                                    let b = obj_ref.borrow();
                                    let x = b.fields.get("left").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    let y = b.fields.get("top").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    let w = b.fields.get("width").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    let h = b.fields.get("height").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    return Ok(Self::make_rectangle(x, y, w, h));
                                }
                                // Computed count properties derived from collections
                                "tabcount" => {
                                    let b = obj_ref.borrow();
                                    if let Some(Value::Collection(coll)) = b.fields.get("tabpages") {
                                        return Ok(Value::Integer(coll.borrow().items.len() as i32));
                                    }
                                    return Ok(Value::Integer(0));
                                }
                                "rowcount" => {
                                    let b = obj_ref.borrow();
                                    if let Some(Value::Collection(coll)) = b.fields.get("rows") {
                                        return Ok(Value::Integer(coll.borrow().items.len() as i32));
                                    }
                                    return Ok(Value::Integer(0));
                                }
                                "columncount" => {
                                    let b = obj_ref.borrow();
                                    if let Some(Value::Collection(coll)) = b.fields.get("columns") {
                                        return Ok(Value::Integer(coll.borrow().items.len() as i32));
                                    }
                                    return Ok(Value::Integer(0));
                                }
                                _ => {}
                            }
                        }
                    }

                    match member_lower.as_str() {
                        "controls" => {
                            // Build the controls collection dynamically from all __is_control fields
                            let b = obj_ref.borrow();
                            let mut ctrl_items = Vec::new();
                            for (_key, val) in &b.fields {
                                if let Value::Object(child_ref) = val {
                                    let child_b = child_ref.borrow();
                                    let is_ctrl = child_b.fields.get("__is_control")
                                        .map(|v| v.as_bool().unwrap_or(false)).unwrap_or(false);
                                    if is_ctrl {
                                        drop(child_b);
                                        ctrl_items.push(val.clone());
                                    }
                                }
                            }
                            // Also include any dynamically-added controls from __controls
                            if let Some(Value::Collection(existing)) = b.fields.get("__controls") {
                                for item in &existing.borrow().items {
                                    // Avoid duplicates: only add if not already in ctrl_items
                                    let already = ctrl_items.iter().any(|ci| {
                                        if let (Value::Object(a), Value::Object(b)) = (ci, item) {
                                            std::rc::Rc::ptr_eq(a, b)
                                        } else { false }
                                    });
                                    if !already {
                                        ctrl_items.push(item.clone());
                                    }
                                }
                            }
                            drop(b);
                            let coll = Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(
                                crate::collections::ArrayList { items: ctrl_items, keys: std::collections::HashMap::new() }
                            )));
                            return Ok(coll);
                        }
                        "components" => return Ok(Value::Nothing),
                        "databindings" => {
                            // Return a DataBindings proxy for this control object
                            let mut flds = std::collections::HashMap::new();
                            flds.insert("__type".to_string(), Value::String("DataBindings".to_string()));
                            flds.insert("__parent".to_string(), Value::Object(obj_ref.clone()));
                            let parent_name_val = {
                                let borrow = obj_ref.borrow();
                                let n = borrow.fields.get("name")
                                    .cloned()
                                    .unwrap_or(Value::String(class_name_str.clone()));
                                n
                            };
                            flds.insert("__parent_name".to_string(), parent_name_val);
                            let proxy = crate::value::ObjectData { class_name: "DataBindings".to_string(), fields: flds };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(proxy))));
                        }
                        "datasource" => {
                            // For DataGridView or any control with a DataSource
                            let borrow = obj_ref.borrow();
                            return Ok(borrow.fields.get("__datasource")
                                .cloned().unwrap_or(Value::Nothing));
                        }
                        _ => {}
                    }
                    
                    // 2. Check if it's a Property Get in the class hierarchy
                    if let Some(prop) = self.find_property(&class_name_str, member.as_str()) {
                         if let Some(body) = &prop.getter {
                             // Construct a temporary FunctionDecl for the getter
                             let func = FunctionDecl {
                                 visibility: prop.visibility,
                                 name: prop.name.clone(),
                                 parameters: prop.parameters.clone(),
                                 return_type: prop.return_type.clone(),
                                 body: body.clone(),
                                 is_async: false,
                                 is_extension: false,
                                 is_overridable: false,
                                 is_overrides: false,
                                 is_must_override: false,
                                 is_shared: false,
                                 is_not_overridable: false,
                             };
                             
                             // Execute Property Get
                             return self.call_user_function(&func, &[], Some(obj_ref.clone()));
                         }
                    }
                }

                if let Value::String(obj_name) = &obj_val {
                    let member_lower = member.as_str().to_lowercase();
                    // DataBindings proxy for string-proxy controls
                    if member_lower == "databindings" {
                        let mut flds = std::collections::HashMap::new();
                        flds.insert("__type".to_string(), Value::String("DataBindings".to_string()));
                        flds.insert("__parent_name".to_string(), Value::String(obj_name.clone()));
                        let proxy = crate::value::ObjectData { class_name: "DataBindings".to_string(), fields: flds };
                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(proxy))));
                    }
                    // DataSource property for string-proxy DataGridView
                    if member_lower == "datasource" {
                        let key = format!("{}.__datasource", obj_name);
                        if let Ok(val) = self.env.get(&key) {
                            return Ok(val);
                        }
                        return Ok(Value::Nothing);
                    }
                    // Rows/Columns for string-proxy DataGridView
                    if member_lower == "rows" || member_lower == "columns" {
                        let key = format!("{}.__{}", obj_name, member_lower);
                        if let Ok(val) = self.env.get(&key) {
                            return Ok(val);
                        }
                        // Auto-create a collection and store it
                        let coll = Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new())));
                        self.env.define_global(&key, coll.clone());
                        return Ok(coll);
                    }
                    // Items/Nodes/TabPages/DropDownItems — auto-create Collection for string-proxy controls
                    if member_lower == "items" || member_lower == "nodes" || member_lower == "tabpages" || member_lower == "dropdownitems" {
                        let key = format!("{}.__{}", obj_name, member_lower);
                        if let Ok(val) = self.env.get(&key) {
                            return Ok(val);
                        }
                        let coll = Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new())));
                        self.env.define_global(&key, coll.clone());
                        return Ok(coll);
                    }
                    // SelectedItems/SelectedRows for string-proxy controls
                    if member_lower == "selecteditems" || member_lower == "selectedrows" {
                        let key = format!("{}.__{}", obj_name, member_lower);
                        if let Ok(val) = self.env.get(&key) {
                            return Ok(val);
                        }
                        let coll = Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new())));
                        self.env.define_global(&key, coll.clone());
                        return Ok(coll);
                    }
                    let key = format!("{}.{}", obj_name, member.as_str());
                    if let Ok(val) = self.env.get(&key) {
                        return Ok(val);
                    }
                    // For string proxy control properties not yet synced, return empty string
                    return Ok(Value::String(String::new()));
                }

                // Accessing a member on Nothing → return Nothing (VB would throw NullReferenceException,
                // but we stay lenient so runtime keeps going).
                if matches!(&obj_val, Value::Nothing) {
                    return Ok(Value::Nothing);
                }
                
                // Fallback to simpler evaluator for non-class objects or if not found
                 evaluate(expr, &self.env)
            }

            Expression::Variable(name) => {
                let var_name = name.as_str();
                let var_lower = var_name.to_lowercase();
                
                // 1. Check local scopes
                if self.env.has_local(var_name) {
                    if let Ok(val) = self.env.get(var_name) {
                        return Ok(val);
                    }
                }

                // 2. Check current object fields (Me.field)
                if let Some(obj_rc) = &self.current_object {
                    let obj = obj_rc.borrow();
                    if let Some(val) = obj.fields.get(&var_lower) {
                        return Ok(val.clone());
                    }
                    drop(obj);
                    
                    // Before auto-creating a field, check global scope first.
                    // Global variables (Console, Math, etc.) should NOT be shadowed
                    // by auto-created Nothing fields on the object.
                    if let Some(val) = self.env.get_global(var_name) {
                        return Ok(val);
                    }
                    // Also check module-level variables
                    if let Some(module) = &self.current_module {
                        let module_key = format!("{}.{}", module, var_name).to_lowercase();
                        if let Ok(val) = self.env.get(&module_key) {
                            return Ok(val);
                        }
                    }
                    
                    // If field not present and not a global, create it as Nothing
                    // to mimic VB's instance fields being default-initialized
                    obj_rc.borrow_mut().fields.insert(var_lower.clone(), Value::Nothing);
                    return Ok(Value::Nothing);
                }

                // 3. Check global scope
                if let Some(val) = self.env.get_global(var_name) {
                    return Ok(val);
                }

                // 4. Try module-level variable if we are in a module
                if let Some(module) = &self.current_module {
                     let module_key = format!("{}.{}", module, var_name).to_lowercase();
                     if let Ok(val) = self.env.get(&module_key) {
                         return Ok(val);
                     }
                }

                // 5. Check imports aliases: `Imports IO = System.IO` → IO resolves to the namespace
                for imp in &self.imports {
                    if let Some(alias) = &imp.alias {
                        if alias.eq_ignore_ascii_case(var_name) {
                            // Try to resolve the path as an env value (namespace object)
                            if let Ok(val) = self.env.get(&imp.path.to_lowercase()) {
                                return Ok(val);
                            }
                        }
                    } else {
                        // Unqualified import: `Imports System.Text` -> `Encoding` resolves to `System.Text.Encoding`
                        let key = format!("{}.{}", imp.path, var_name).to_lowercase();
                        if let Ok(val) = self.env.get(&key) {
                            return Ok(val);
                        }
                    }
                }

                // 6. Fallback: implicit function call without parentheses (e.g. "Now", "Date")
                match self.call_procedure(name, &[]) {
                    Ok(val) => return Ok(val),
                    Err(_) => return Err(RuntimeError::UndefinedVariable(var_name.to_string())),
                }
            }


            // For simple expressions, use the standalone evaluator
            _ => evaluate(expr, &self.env),
        }
    }



    pub fn call_procedure(&mut self, name: &Identifier, args: &[Expression]) -> Result<Value, RuntimeError> {
        let name_str = name.as_str().to_lowercase();


        // First check if it's actually an array access (arrays and functions use same syntax in VB)
        // Check local variable shadowing first
        if let Ok(val) = self.env.get(name.as_str()) {
             if let Value::Array(_) = val {
                 // It's an array access
                 if args.len() == 1 {
                     let index = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                     return val.get_array_element(index);
                 }
             }
             if let Value::Collection(_) = val {
                 // It's a collection item access
                 if args.len() == 1 {
                     let index = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                     return val.get_array_element(index);
                 }
             }
             if let Value::Lambda { .. } = &val {
                 // It's a lambda/delegate call (Sub or Function lambda invoked as statement)
                 let arg_values: Result<Vec<_>, _> = args.iter().map(|e| self.evaluate_expr(e)).collect();
                 return self.call_lambda(val, &arg_values?);
             }
        }
        
        // Check implicit object array/field access
        let current_obj_opt = self.current_object.clone();
        if let Some(obj_rc) = current_obj_opt {
             let obj = obj_rc.borrow();
             if let Some(val) = obj.fields.get(name.as_str()) {
                 if let Value::Array(_) = val {
                     // Array access on field
                     if args.len() == 1 {
                         let index = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                         return val.get_array_element(index);
                     }
                 }
             }
        }




        // If we're inside an object context, try resolving as a class method
        if let Some(obj_rc) = &self.current_object {
            let class_name = obj_rc.borrow().class_name.clone();
            if let Some(method) = self.find_method(&class_name, name.as_str()) {
                match method {
                    vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                        self.call_user_sub_exprs(&s, args, Some(obj_rc.clone()))?;
                        return Ok(Value::Nothing);
                    }
                    vybe_parser::ast::decl::MethodDecl::Function(f) => {
                        return self.call_user_function_exprs(&f, args, Some(obj_rc.clone()));
                    }
                }
            }
        }

        // Check user-defined functions with scope resolution
        // Try 1: Look for qualified name (as-is)
        if let Some(func) = self.functions.get(&name_str).cloned() {
            return self.call_user_function_exprs(&func, args, None);
        }
        // Try 2: If in a module, try current_module.name
        if let Some(module) = &self.current_module {
            let qualified = format!("{}.{}", module.to_lowercase(), name_str);
            if let Some(func) = self.functions.get(&qualified).cloned() {
                return self.call_user_function_exprs(&func, args, None);
            }
        }
        // Try 3: Check imports-qualified names
        for imp in &self.imports {
            if imp.alias.is_none() {
                let qualified = format!("{}.{}", imp.path.to_lowercase(), name_str);
                if let Some(func) = self.functions.get(&qualified).cloned() {
                    return self.call_user_function_exprs(&func, args, None);
                }
            }
        }
        // Try 4: Global search - look for any function with matching unqualified name
        // This makes BAS module functions globally accessible
        for (key, func) in &self.functions {
            if key.ends_with(&format!(".{}", name_str)) || key == &name_str {
                let func_clone = func.clone();
                return self.call_user_function_exprs(&func_clone, args, None);
            }
        }

        // Also check if it's a Sub (in VB, both Subs and Functions can be called with parens)
        // Try 1: Look for qualified name (as-is)
        if let Some(sub) = self.subs.get(&name_str).cloned() {
            return self.call_user_sub_exprs(&sub, args, None);
        }
        // Try 2: If in a module, try current_module.name
        if let Some(module) = &self.current_module {
            let qualified = format!("{}.{}", module.to_lowercase(), name_str);
            if let Some(sub) = self.subs.get(&qualified).cloned() {
                return self.call_user_sub_exprs(&sub, args, None);
            }
        }
        // Try 3: Check imports-qualified names
        for imp in &self.imports {
            if imp.alias.is_none() {
                let qualified = format!("{}.{}", imp.path.to_lowercase(), name_str);
                if let Some(sub) = self.subs.get(&qualified).cloned() {
                    return self.call_user_sub_exprs(&sub, args, None);
                }
            }
        }
        // Try 4: Global search for subs
        for (key, sub) in &self.subs {
            if key.ends_with(&format!(".{}", name_str)) || key == &name_str {
                return self.call_user_sub_exprs(&sub.clone(), args, None);
            }
        }


    
    // Fallback: Check built-in functions
    if let Ok(arg_values) = args.iter().map(|e| self.evaluate_expr(e)).collect::<Result<Vec<_>, _>>() {
        // Try standard library first
        if let Ok(val) = crate::std_lib::call_builtin(&name_str, &arg_values) {
            return Ok(val);
        }

        match name_str.as_str() {
            "msgbox" => {
                let msg = if arg_values.is_empty() {
                    "".to_string()
                } else {
                    arg_values[0].as_string()
                };
                let buttons = if arg_values.len() >= 2 {
                    match &arg_values[1] {
                        Value::Integer(i) => *i,
                        _ => 0,
                    }
                } else {
                    0
                };
                let title = if arg_values.len() >= 3 {
                    arg_values[2].as_string()
                } else {
                    "vybe Basic".to_string()
                };
                // Use native OS dialog that blocks synchronously
                let result = crate::builtins::info_fns::show_native_msgbox(&msg, &title, buttons);
                return Ok(Value::Integer(result));
            }
            "inputbox" => return inputbox_fn(&arg_values),

            // String functions
            "len" => return len_fn(&arg_values),
            "left" => return left_fn(&arg_values),
            "right" => return right_fn(&arg_values),
            "mid" | "mid$" => return mid_fn(&arg_values),
            "ucase" | "ucase$" => return ucase_fn(&arg_values),
            "lcase" | "lcase$" => return lcase_fn(&arg_values),
            "trim" | "trim$" => return trim_fn(&arg_values),
            "ltrim" | "ltrim$" => return ltrim_fn(&arg_values),
            "rtrim" | "rtrim$" => return rtrim_fn(&arg_values),
            "instr" => return instr_fn(&arg_values),
            "instrrev" => return instrrev_fn(&arg_values),
            "replace" => return replace_fn(&arg_values),
            "chr" | "chr$" => return chr_fn(&arg_values),
            "asc" => return asc_fn(&arg_values),
            "split" => return split_fn(&arg_values),
            "join" => return join_fn(&arg_values),
            "strreverse" => return strreverse_fn(&arg_values),
            "space" | "space$" | "spc" => return space_fn(&arg_values),
            "string" | "string$" => return string_repeat_fn(&arg_values),
            "strcomp" => return strcomp_fn(&arg_values),
            "format" | "format$" => return format_fn(&arg_values),

            // Array functions
            "ubound" => return ubound_fn(&arg_values),
            "lbound" => return lbound_fn(&arg_values),
            "array" => return array_fn(&arg_values),
            "erase" => return erase_fn(&arg_values),

            // Conversion functions
            "cstr" => return cstr_fn(&arg_values),
            "cint" => return cint_fn(&arg_values),
            "cdbl" => return cdbl_fn(&arg_values),
            "cbool" => return cbool_fn(&arg_values),
            "clng" => return clng_fn(&arg_values),
            "csng" => return csng_fn(&arg_values),
            "cbyte" => return cbyte_fn(&arg_values),
            "cchar" => return cchar_fn(&arg_values),
            "val" => return val_fn(&arg_values),
            "str" | "str$" => return str_fn(&arg_values),
            "hex" | "hex$" => return hex_fn(&arg_values),
            "oct" | "oct$" => return oct_fn(&arg_values),

            // Math functions
            "abs" => return abs_fn(&arg_values),
            "int" => return int_fn(&arg_values),
            "fix" => return fix_fn(&arg_values),
            "sgn" => return sgn_fn(&arg_values),
            "sqr" => return sqr_fn(&arg_values),
            "sqrt" => return sqr_fn(&arg_values), // .NET alias
            "rnd" => return rnd_fn(&arg_values),
            "round" => return round_fn(&arg_values),
            "log" => return log_fn(&arg_values),
            "exp" => return exp_fn(&arg_values),
            "sin" => return sin_fn(&arg_values),
            "cos" => return cos_fn(&arg_values),
            "tan" => return tan_fn(&arg_values),
            "atn" | "atan" => return atn_fn(&arg_values),

            // Type info functions
            "isnumeric" => return isnumeric_fn(&arg_values),
            "isarray" => return isarray_fn(&arg_values),
            "isnothing" => return isnothing_fn(&arg_values),
            "isdate" => return isdate_fn(&arg_values),
            "isnull" => return isnull_fn(&arg_values),
            "typename" => return typename_fn(&arg_values),
            "vartype" => return vartype_fn(&arg_values),

            // Utility functions
            "iif" => return iif_fn(&arg_values),
            "choose" => return choose_fn(&arg_values),
            "switch" => return switch_fn(&arg_values),

            // Date/Time functions
            "now" => return now_fn(&arg_values),
            "date" | "today" => return date_fn(&arg_values),
            "time" | "timeofday" => return time_fn(&arg_values),
            "year" => return year_fn(&arg_values),
            "month" => return month_fn(&arg_values),
            "day" => return day_fn(&arg_values),
            "hour" => return hour_fn(&arg_values),
            "minute" => return minute_fn(&arg_values),
            "second" => return second_fn(&arg_values),
            "timer" => return timer_fn(&arg_values),

            // Additional math functions
            "max" | "math.max" => return max_fn(&arg_values),
            "min" | "math.min" => return min_fn(&arg_values),
            "ceiling" | "math.ceiling" => return ceiling_fn(&arg_values),
            "floor" | "math.floor" => return floor_fn(&arg_values),
            "pow" | "math.pow" => return pow_fn(&arg_values),
            "randomize" => return randomize_fn(&arg_values),
            "atan2" | "math.atan2" => return atan2_fn(&arg_values),

            // Additional conversion functions

            "cdate" => return cdate_fn(&arg_values),
            "cdec" => return cdec_fn(&arg_values),
            "ccur" => return ccur_fn(&arg_values),
            "cvar" => return cvar_fn(&arg_values),
            "formatnumber" => return formatnumber_fn(&arg_values),
            "formatcurrency" => return formatcurrency_fn(&arg_values),
            "formatpercent" => return formatpercent_fn(&arg_values),

            // Misc utility functions
            "doevents" => return doevents_fn(&arg_values),
            "isnullorempty" | "string.isnullorempty" => return isnullorempty_fn(&arg_values),
            "strdup" => return strdup_fn(&arg_values),

            // Color functions (moved from misc to conversion_fns)
            "rgb" => return rgb_fn(&arg_values),
            "qbcolor" => return qbcolor_fn(&arg_values),

            // String functions additions
            "strconv" => return strconv_fn(&arg_values),
            "lset" | "lset$" => return lset_fn(&arg_values),
            "rset" | "rset$" => return rset_fn(&arg_values),
            "filter" => return filter_fn(&arg_values),
            "formatdatetime" => return formatdatetime_fn(&arg_values),

            // Conversion functions additions
            "cobj" => return cobj_fn(&arg_values),
            "cshort" => return cshort_fn(&arg_values),
            "cushort" => return cushort_fn(&arg_values),
            "cuint" => return cuint_fn(&arg_values),
            "culng" => return culng_fn(&arg_values),
            "csbyte" => return csbyte_fn(&arg_values),
            "ascw" => return ascw_fn(&arg_values),
            "chrw" | "chrw$" => return chrw_fn(&arg_values),
            // DateTime functions additions
            "dateadd" => return dateadd_fn(&arg_values),
            "datediff" => return datediff_fn(&arg_values),
            "datepart" => return datepart_fn(&arg_values),
            "dateserial" => return dateserial_fn(&arg_values),
            "timeserial" => return timeserial_fn(&arg_values),
            "datevalue" => return datevalue_fn(&arg_values),
            "timevalue" => return timevalue_fn(&arg_values),
            "monthname" => return monthname_fn(&arg_values),
            "weekdayname" => return weekdayname_fn(&arg_values),
            "weekday" => return weekday_fn(&arg_values),

            // Info functions additions
            "isempty" => return isempty_fn(&arg_values),
            "isobject" => return isobject_fn(&arg_values),
            "iserror" => return iserror_fn(&arg_values),
            "isdbnull" => return isdbnull_fn(&arg_values),

            // File functions
            "dir" | "dir$" => return dir_fn(&arg_values),
            "filecopy" => return filecopy_fn(&arg_values),
            "kill" => return kill_fn(&arg_values),
            "name" => return name_fn(&arg_values),
            "getattr" => return getattr_fn(&arg_values),
            "setattr" => return setattr_fn(&arg_values),
            "filedatetime" => return filedatetime_fn(&arg_values),
            "filelen" => return filelen_fn(&arg_values),
            "curdir" | "curdir$" => return curdir_fn(&arg_values),
            "chdir" => return chdir_fn(&arg_values),
            "mkdir" => return mkdir_fn(&arg_values),
            "rmdir" => return rmdir_fn(&arg_values),
            "freefile" => return freefile_fn(&arg_values),
            "fileexists" | "file.exists" => return file_exists_fn(&arg_values),
            
            // VB6 File Handle functions
            "eof" => return eof_fn(&arg_values),
            "lof" => return lof_fn(&arg_values),
            "loc" => return loc_fn(&arg_values),
            "fileattr" => return fileattr_fn(&arg_values),
            "input" => return input_fn(&arg_values),
            "inputb" => return inputb_fn(&arg_values),
            "open" => return open_file_fn(&arg_values),
            "close" => return close_file_fn(&arg_values),
            "print" => return print_file_fn(&arg_values),
            "write" => return write_file_fn(&arg_values),
            "lineinput" => return line_input_fn(&arg_values),
            "seek" => return seek_file_fn(&arg_values),
            "get" => return get_file_fn(&arg_values),
            "put" => return put_file_fn(&arg_values),
            
            // Image functions
            "loadpicture" => return loadpicture_fn(&arg_values),
            "savepicture" => return savepicture_fn(&arg_values),

            // Interaction functions
            "beep" => return beep_fn(&arg_values),
            "shell" => return shell_fn(&arg_values),
            "environ" | "environ$" => return environ_fn(&arg_values),
            "command" | "command$" => {
                return Ok(Value::String(self.command_line_args.join(" ")));
            }
            "sendkeys" => return sendkeys_fn(&arg_values),
            "appactivate" => return appactivate_fn(&arg_values),
            "load" => return load_fn(&arg_values),
            "unload" => return unload_fn(&arg_values),
            "app" => return app_fn(&arg_values),
            "screen" => return screen_fn(&arg_values),
            "clipboard" => return clipboard_fn(&arg_values),
            "forms" => return forms_fn(&arg_values),

            // System.Text functions
            "stringbuilder" => return stringbuilder_new_fn(&arg_values),
            
            // Encoding functions
            "encoding.ascii.getbytes" => return encoding_ascii_getbytes_fn(&arg_values),
            "encoding.ascii.getstring" => return encoding_ascii_getstring_fn(&arg_values),
            "encoding.utf8.getbytes" => return encoding_utf8_getbytes_fn(&arg_values),
            "encoding.utf8.getstring" => return encoding_utf8_getstring_fn(&arg_values),
            "encoding.unicode.getbytes" => return encoding_unicode_getbytes_fn(&arg_values),
            "encoding.unicode.getstring" => return encoding_unicode_getstring_fn(&arg_values),
            "encoding.default.getbytes" => return encoding_default_getbytes_fn(&arg_values),
            "encoding.default.getstring" => return encoding_default_getstring_fn(&arg_values),
            "encoding.getencoding" => return encoding_getencoding_fn(&arg_values),
            "encoding.convert" => return encoding_convert_fn(&arg_values),
            
            // Regex functions
            "regex.ismatch" => return regex_ismatch_fn(&arg_values),
            "regex.match" => return regex_match_fn(&arg_values),
            "regex.matches" => return regex_matches_fn(&arg_values),
            "regex.replace" => return regex_replace_fn(&arg_values),
            "regex.split" => return regex_split_fn(&arg_values),
            
            // JSON functions
            "jsonserializer.serialize" | "json.serialize" => return json_serialize_fn(&arg_values),
            "jsonserializer.deserialize" | "json.deserialize" => return json_deserialize_fn(&arg_values),
            
            // XML functions
            "xdocument.parse" | "xml.parse" => return crate::builtins::xml::xdocument_parse(&arg_values),
            "xdocument.load" | "xml.load" => return crate::builtins::xml::xdocument_load(&arg_values),
            "xdocument.save" | "xml.save" => return xml_save_fn(&arg_values),
            "xelement.parse" => return crate::builtins::xml::xelement_parse(&arg_values),

            // Financial functions
            "pmt" => return pmt_fn(&arg_values),
            "fv" => return fv_fn(&arg_values),
            "pv" => return pv_fn(&arg_values),
            "nper" => return nper_fn(&arg_values),
            "rate" => return rate_fn(&arg_values),
            "ipmt" => return ipmt_fn(&arg_values),
            "ppmt" => return ppmt_fn(&arg_values),
            "ddb" => return ddb_fn(&arg_values),
            "sln" => return sln_fn(&arg_values),
            "syd" => return syd_fn(&arg_values),

            // Debug/Output
            "debug.print" => {
                let msg = arg_values.iter().map(|v| v.as_string()).collect::<Vec<_>>().join(" ");
                self.send_debug_output(format!("{}\n", msg));
                return Ok(Value::Nothing);
            }
            "console.writeline" => {
                return self.dispatch_console_method("writeline", &arg_values);
            }
            "console.write" => {
                return self.dispatch_console_method("write", &arg_values);
            }
            "console.readline" => {
                return self.dispatch_console_method("readline", &arg_values);
            }
            "console.read" => {
                return self.dispatch_console_method("read", &arg_values);
            }
            "console.clear" => {
                return self.dispatch_console_method("clear", &arg_values);
            }
            "console.resetcolor" => {
                return self.dispatch_console_method("resetcolor", &arg_values);
            }
            "console.beep" => {
                return self.dispatch_console_method("beep", &arg_values);
            }
            "console.setcursorposition" => {
                return self.dispatch_console_method("setcursorposition", &arg_values);
            }

            "application.run" | "system.windows.forms.application.run" => {
                return self.dispatch_application_method("run", &arg_values);
            }
            "application.exit" | "system.windows.forms.application.exit"
            | "my.application.exit" => {
                return self.dispatch_application_method("exit", &arg_values);
            }
            "my.application.doevents" => {
                return self.dispatch_application_method("doevents", &arg_values);
            }

            "callbyname" | "microsoft.visualbasic.interaction.callbyname" => {
                let obj_val = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                let member = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                let call_type = arg_values.get(2).map(|v| v.as_integer().unwrap_or(1)).unwrap_or(1);
                let member_lower = member.to_lowercase();
                
                if let Value::Object(obj_ref) = &obj_val {
                    match call_type {
                        2 | 8 => {
                            // Get property or field
                            let obj_data = obj_ref.borrow();
                            if let Some(val) = obj_data.fields.get(&member_lower) {
                                return Ok(val.clone());
                            }
                            return Ok(Value::Nothing);
                        }
                        4 => {
                            // Set property or field
                            let set_val = arg_values.get(3).cloned().unwrap_or(Value::Nothing);
                            obj_ref.borrow_mut().fields.insert(member_lower, set_val);
                            return Ok(Value::Nothing);
                        }
                        _ => {
                            // Method call (1)
                            // CallByName(obj, "MethodName", 1, arg1, arg2...)
                            // Arguments start at index 3
                            let method_args = if arg_values.len() > 3 {
                                &arg_values[3..]
                            } else {
                                &[]
                            };

                            let class_name = obj_ref.borrow().class_name.clone();
                            if let Some(method_decl) = self.find_method(&class_name, &member) {
                                match method_decl {
                                    vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                                        self.call_user_sub(&s, method_args, Some(obj_ref.clone()))?;
                                        return Ok(Value::Nothing);
                                    }
                                    vybe_parser::ast::decl::MethodDecl::Function(f) => {
                                        let result = self.call_user_function(&f, method_args, Some(obj_ref.clone()))?;
                                        return Ok(result);
                                    }
                                }
                            } else {
                                return Err(RuntimeError::UndefinedFunction(format!("Method '{}' not found in class '{}'", member, class_name)));
                            }
                        }
                    }
                }
                return Ok(Value::Nothing);
            }

            _ => {}
        }
    }

    Err(RuntimeError::UndefinedFunction(name.as_str().to_string()))
    }

    fn call_function(&mut self, name: &Identifier, args: &[Expression]) -> Result<Value, RuntimeError> {
        self.call_procedure(name, args)
    }

    fn call_method(&mut self, obj: &Expression, method: &Identifier, args: &[Expression]) -> Result<Value, RuntimeError> {
        let method_name = method.as_str().to_lowercase();

        // ── Static Class Dispatch ───────────────────────────────────────
        if let Expression::Variable(name) = obj {
            let class_name = name.as_str().to_lowercase();
            match class_name.as_str() {
                "math" | "system.math" => {
                    let arg_values: Vec<Value> = args.iter().map(|a| self.evaluate_expr(a)).collect::<Result<_,_>>()?;
                    use crate::builtins::math_fns::*;
                    match method_name.as_str() {
                        "abs" => return abs_fn(&arg_values),
                        "ceiling" => return ceiling_fn(&arg_values),
                        "cos" => return cos_fn(&arg_values),
                        "exp" => return exp_fn(&arg_values),
                        "floor" => return floor_fn(&arg_values),
                        "log" => return log_fn(&arg_values),
                        "max" => return max_fn(&arg_values),
                        "min" => return min_fn(&arg_values),
                        "pow" => return pow_fn(&arg_values),
                        "round" => return round_fn(&arg_values),
                        "sign" => return sgn_fn(&arg_values),
                        "sin" => return sin_fn(&arg_values),
                        "sqrt" => return sqr_fn(&arg_values),
                        "tan" => return tan_fn(&arg_values),
                        "truncate" => return fix_fn(&arg_values), // Fix truncates
                        "atan" => return atn_fn(&arg_values),
                        "atan2" => return atan2_fn(&arg_values),
                        _ => return Err(RuntimeError::UndefinedFunction(format!("Math.{}", method_name))),
                    }
                }
                "file" | "system.io.file" => {
                    let arg_values: Vec<Value> = args.iter().map(|a| self.evaluate_expr(a)).collect::<Result<_,_>>()?;
                    use crate::builtins::file_fns::*;
                    match method_name.as_str() {
                        "readalltext" => return file_readalltext_fn(&arg_values),
                        "writealltext" => return file_writealltext_fn(&arg_values),
                        "appendalltext" => return file_appendalltext_fn(&arg_values),
                        "readalllines" => return file_readalllines_fn(&arg_values),
                        "exists" => return file_exists_fn(&arg_values),
                        "delete" => return file_delete_fn(&arg_values),
                        "copy" => return file_copy_fn(&arg_values),
                        "move" => return file_move_fn(&arg_values),
                        _ => return Err(RuntimeError::UndefinedFunction(format!("File.{}", method_name))),
                    }
                }
                "directory" | "system.io.directory" => {
                    let arg_values: Vec<Value> = args.iter().map(|a| self.evaluate_expr(a)).collect::<Result<_,_>>()?;
                    use crate::builtins::file_fns::*;
                    match method_name.as_str() {
                        "exists" => return directory_exists_fn(&arg_values),
                        "createdirectory" => return directory_createdirectory_fn(&arg_values),
                        "delete" => return directory_delete_fn(&arg_values),
                        "getfiles" => return directory_getfiles_fn(&arg_values),
                        "getdirectories" => return directory_getdirectories_fn(&arg_values),
                        "getcurrentdirectory" => return directory_getcurrentdirectory_fn(),
                        _ => return Err(RuntimeError::UndefinedFunction(format!("Directory.{}", method_name))),
                    }
                }
                "path" | "system.io.path" => {
                    let arg_values: Vec<Value> = args.iter().map(|a| self.evaluate_expr(a)).collect::<Result<_,_>>()?;
                    use crate::builtins::file_fns::*;
                    match method_name.as_str() {
                        "combine" => return path_combine_fn(&arg_values),
                        "getfilename" => return path_getfilename_fn(&arg_values),
                        "getdirectoryname" => return path_getdirectoryname_fn(&arg_values),
                        "getextension" => return path_getextension_fn(&arg_values),
                        "changeextension" => return path_changeextension_fn(&arg_values),
                        _ => return Err(RuntimeError::UndefinedFunction(format!("Path.{}", method_name))),
                    }
                }
                "md5" | "system.security.cryptography.md5" => {
                    match method_name.as_str() {
                        "create" => {
                            let mut fields = HashMap::new();
                            fields.insert("__type".to_string(), Value::String("MD5".to_string()));
                            let obj = crate::value::ObjectData { class_name: "MD5".to_string(), fields };
                            return Ok(Value::Object(Rc::new(RefCell::new(obj))));
                        }
                        _ => return Err(RuntimeError::UndefinedFunction(format!("MD5.{}", method_name))),
                    }
                }
                "sha256" | "system.security.cryptography.sha256" => {
                    match method_name.as_str() {
                        "create" => {
                            let mut fields = HashMap::new();
                            fields.insert("__type".to_string(), Value::String("SHA256".to_string()));
                            let obj = crate::value::ObjectData { class_name: "SHA256".to_string(), fields };
                            return Ok(Value::Object(Rc::new(RefCell::new(obj))));
                        }
                        _ => return Err(RuntimeError::UndefinedFunction(format!("SHA256.{}", method_name))),
                    }
                }
                "xdocument" | "system.xml.linq.xdocument" => {
                    let arg_values: Vec<Value> = args.iter().map(|a| self.evaluate_expr(a)).collect::<Result<_,_>>()?;
                    match method_name.as_str() {
                        "parse" => return crate::builtins::xml::xdocument_parse(&arg_values),
                        "load" => return crate::builtins::xml::xdocument_load(&arg_values),
                        _ => return Err(RuntimeError::UndefinedFunction(format!("XDocument.{}", method_name))),
                    }
                }
                "xelement" | "system.xml.linq.xelement" => {
                    let arg_values: Vec<Value> = args.iter().map(|a| self.evaluate_expr(a)).collect::<Result<_,_>>()?;
                    match method_name.as_str() {
                        "parse" => return crate::builtins::xml::xelement_parse(&arg_values),
                        _ => return Err(RuntimeError::UndefinedFunction(format!("XElement.{}", method_name))),
                    }
                }
                _ => {}
            }
        }

        // ── MyBase dispatch ─────────────────────────────────────────────
        // MyBase.Method() dispatches to the parent class's method on the
        // same object instance (Me).
        let is_mybase = matches!(obj, Expression::MyBase);
        if is_mybase {
            if let Some(obj_rc) = &self.current_object {
                let obj_clone = obj_rc.clone();
                let class_name = obj_clone.borrow().class_name.clone();
                let arg_values: Vec<Value> = args.iter()
                    .map(|a| self.evaluate_expr(a))
                    .collect::<Result<Vec<_>, _>>()?;
                if let Some(base_method) = self.find_method_in_base(&class_name, &method_name) {
                    match base_method {
                        vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                            self.call_user_sub(&s, &arg_values, Some(obj_clone))?;
                            return Ok(Value::Nothing);
                        }
                        vybe_parser::ast::decl::MethodDecl::Function(f) => {
                            return self.call_user_function(&f, &arg_values, Some(obj_clone));
                        }
                    }
                }
                return Err(RuntimeError::UndefinedFunction(
                    format!("MyBase.{}", method.as_str()),
                ));
            }
            return Err(RuntimeError::Custom("'MyBase' used outside of object context".to_string()));
        }

        // Handle WinForms designer no-op methods
        match method_name.as_str() {
            "suspendlayout" | "resumelayout" | "performlayout" => return Ok(Value::Nothing),
            _ => {}
        }

        // Handle Me.Controls("Button1") default indexer — MethodCall(Me, "Controls", ["Button1"])
        if method_name == "controls" && !args.is_empty() {
            // Build the Controls collection from the parent object, then index by string key or integer
            let parent_val = self.evaluate_expr(obj)?;
            if let Value::Object(parent_ref) = &parent_val {
                // Build the controls list dynamically
                let pb = parent_ref.borrow();
                let mut controls: Vec<Value> = Vec::new();
                for (_, v) in &pb.fields {
                    if let Value::Object(child) = v {
                        let is_ctrl = child.borrow().fields.get("__is_control")
                            .map(|v| v.as_bool().unwrap_or(false)).unwrap_or(false);
                        if is_ctrl {
                            controls.push(v.clone());
                        }
                    }
                }
                if let Some(Value::Collection(coll)) = pb.fields.get("__controls") {
                    for item in &coll.borrow().items {
                        controls.push(item.clone());
                    }
                }
                drop(pb);

                let idx_val = self.evaluate_expr(&args[0])?;
                match &idx_val {
                    Value::String(key) => {
                        // Search by name
                        for ctrl in &controls {
                            if let Value::Object(o) = ctrl {
                                let matches = o.borrow().fields.get("name")
                                    .map(|v| v.as_string().eq_ignore_ascii_case(key))
                                    .unwrap_or(false);
                                if matches {
                                    return Ok(ctrl.clone());
                                }
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    Value::Integer(i) => {
                        let idx = *i as usize;
                        return Ok(controls.get(idx).cloned().unwrap_or(Value::Nothing));
                    }
                    _ => {
                        if let Ok(i) = idx_val.as_integer() {
                            let idx = i as usize;
                            return Ok(controls.get(idx).cloned().unwrap_or(Value::Nothing));
                        }
                        return Ok(Value::Nothing);
                    }
                }
            }
        }

        // Handle Controls.* methods — full ControlCollection API
        if let Expression::MemberAccess(parent_expr, member) = obj {
            if member.as_str().eq_ignore_ascii_case("Controls") {
                // Extract parent control name from parent_expr
                // Me.pnlRoot.Controls → parent_name = "pnlRoot"
                // Me.Controls → parent_name = "" (form-level)
                let parent_ctrl_name = match parent_expr.as_ref() {
                    Expression::MemberAccess(_, id) => {
                        let name = id.as_str();
                        // "Me" means form-level
                        if name.eq_ignore_ascii_case("me") { String::new() } else { name.to_string() }
                    }
                    _ => String::new(), // Variable("Me") or other → form-level
                };

                match method_name.as_str() {
                    "add" => {
                        if !args.is_empty() {
                            let ctrl_val = self.evaluate_expr(&args[0])?;
                            if let Value::Object(ctrl_obj) = &ctrl_val {
                                let b = ctrl_obj.borrow();
                                let ctrl_type = b.fields.get("__type").map(|v| v.as_string()).unwrap_or_default();
                                let ctrl_name = b.fields.get("name").map(|v| v.as_string()).unwrap_or_default();
                                let left = b.fields.get("left").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                let top = b.fields.get("top").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                let width = b.fields.get("width").and_then(|v| v.as_integer().ok()).unwrap_or(100);
                                let height = b.fields.get("height").and_then(|v| v.as_integer().ok()).unwrap_or(30);
                                drop(b);
                                if !ctrl_name.is_empty() {
                                    self.side_effects.push_back(crate::RuntimeSideEffect::AddControl {
                                        form_name: String::new(),
                                        control_name: ctrl_name.clone(),
                                        control_type: ctrl_type,
                                        left, top, width, height,
                                        parent_name: parent_ctrl_name.clone(),
                                    });
                                }
                                // Also store the control as a field on the parent for Controls iteration
                                if let Ok(parent_val) = self.evaluate_expr(parent_expr) {
                                    if let Value::Object(parent_ref) = &parent_val {
                                        let mut pb = parent_ref.borrow_mut();
                                        // Store as named field so Controls accessor finds it
                                        if !ctrl_name.is_empty() {
                                            pb.fields.insert(ctrl_name.to_lowercase(), ctrl_val.clone());
                                        }
                                        // Also track in __controls for dynamically added controls
                                        if let Some(Value::Collection(coll)) = pb.fields.get("__controls") {
                                            coll.borrow_mut().items.push(ctrl_val.clone());
                                        } else {
                                            let mut al = crate::collections::ArrayList::new();
                                            al.items.push(ctrl_val.clone());
                                            pb.fields.insert("__controls".to_string(),
                                                Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(al))));
                                        }
                                    }
                                }
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    "addrange" => {
                        // Controls.AddRange(controls()) — add multiple controls
                        if !args.is_empty() {
                            let val = self.evaluate_expr(&args[0])?;
                            let items = match val {
                                Value::Array(arr) => arr,
                                Value::Collection(c) => c.borrow().items.clone(),
                                _ => vec![val],
                            };
                            for ctrl_val in items {
                                if let Value::Object(ctrl_obj) = &ctrl_val {
                                    let b = ctrl_obj.borrow();
                                    let ctrl_type = b.fields.get("__type").map(|v| v.as_string()).unwrap_or_default();
                                    let ctrl_name = b.fields.get("name").map(|v| v.as_string()).unwrap_or_default();
                                    let left = b.fields.get("left").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    let top = b.fields.get("top").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    let width = b.fields.get("width").and_then(|v| v.as_integer().ok()).unwrap_or(100);
                                    let height = b.fields.get("height").and_then(|v| v.as_integer().ok()).unwrap_or(30);
                                    drop(b);
                                    if !ctrl_name.is_empty() {
                                        self.side_effects.push_back(crate::RuntimeSideEffect::AddControl {
                                            form_name: String::new(),
                                            control_name: ctrl_name.clone(),
                                            control_type: ctrl_type,
                                            left, top, width, height,
                                            parent_name: parent_ctrl_name.clone(),
                                        });
                                    }
                                    if let Ok(parent_val) = self.evaluate_expr(parent_expr) {
                                        if let Value::Object(parent_ref) = &parent_val {
                                            let mut pb = parent_ref.borrow_mut();
                                            if !ctrl_name.is_empty() {
                                                pb.fields.insert(ctrl_name.to_lowercase(), ctrl_val.clone());
                                            }
                                            if let Some(Value::Collection(coll)) = pb.fields.get("__controls") {
                                                coll.borrow_mut().items.push(ctrl_val.clone());
                                            } else {
                                                let mut al = crate::collections::ArrayList::new();
                                                al.items.push(ctrl_val.clone());
                                                pb.fields.insert("__controls".to_string(),
                                                    Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(al))));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    "clear" => {
                        // Controls.Clear() — remove all controls
                        if let Ok(parent_val) = self.evaluate_expr(parent_expr) {
                            if let Value::Object(parent_ref) = &parent_val {
                                let mut pb = parent_ref.borrow_mut();
                                // Remove all __is_control fields
                                let ctrl_keys: Vec<String> = pb.fields.iter()
                                    .filter(|(_, v)| {
                                        if let Value::Object(o) = v {
                                            o.borrow().fields.get("__is_control")
                                                .map(|v| v.as_bool().unwrap_or(false)).unwrap_or(false)
                                        } else { false }
                                    })
                                    .map(|(k, _)| k.clone())
                                    .collect();
                                for key in &ctrl_keys {
                                    // Emit property change to hide from UI
                                    if let Some(Value::Object(ctrl_ref)) = pb.fields.get(key) {
                                        let ctrl_name = ctrl_ref.borrow().fields.get("name")
                                            .map(|v| v.as_string()).unwrap_or_default();
                                        if !ctrl_name.is_empty() {
                                            self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                                                object: ctrl_name,
                                                property: "Visible".to_string(),
                                                value: Value::Boolean(false),
                                            });
                                        }
                                    }
                                    pb.fields.remove(key);
                                }
                                // Clear __controls too
                                if let Some(Value::Collection(coll)) = pb.fields.get("__controls") {
                                    coll.borrow_mut().clear();
                                }
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    "remove" => {
                        // Controls.Remove(ctrl) — remove a specific control
                        if !args.is_empty() {
                            let ctrl_val = self.evaluate_expr(&args[0])?;
                            if let Value::Object(ctrl_obj) = &ctrl_val {
                                let ctrl_name = ctrl_obj.borrow().fields.get("name")
                                    .map(|v| v.as_string()).unwrap_or_default();
                                if !ctrl_name.is_empty() {
                                    self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                                        object: ctrl_name.clone(),
                                        property: "Visible".to_string(),
                                        value: Value::Boolean(false),
                                    });
                                    if let Ok(parent_val) = self.evaluate_expr(parent_expr) {
                                        if let Value::Object(parent_ref) = &parent_val {
                                            parent_ref.borrow_mut().fields.remove(&ctrl_name.to_lowercase());
                                        }
                                    }
                                }
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    "removeat" => {
                        // Controls.RemoveAt(index) — remove by index
                        if !args.is_empty() {
                            let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                            // Get the controls list, find the control at index, remove it
                            let controls_val = self.evaluate_expr(&Expression::MemberAccess(
                                parent_expr.clone(), member.clone()))?;
                            if let Value::Collection(coll) = &controls_val {
                                let items = coll.borrow().items.clone();
                                if idx < items.len() {
                                    if let Value::Object(ctrl_obj) = &items[idx] {
                                        let ctrl_name = ctrl_obj.borrow().fields.get("name")
                                            .map(|v| v.as_string()).unwrap_or_default();
                                        if !ctrl_name.is_empty() {
                                            self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                                                object: ctrl_name.clone(),
                                                property: "Visible".to_string(),
                                                value: Value::Boolean(false),
                                            });
                                            if let Ok(parent_val) = self.evaluate_expr(parent_expr) {
                                                if let Value::Object(parent_ref) = &parent_val {
                                                    parent_ref.borrow_mut().fields.remove(&ctrl_name.to_lowercase());
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    "removebykey" => {
                        // Controls.RemoveByKey(key) — remove by Name string
                        if !args.is_empty() {
                            let key = self.evaluate_expr(&args[0])?.as_string();
                            self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                                object: key.clone(),
                                property: "Visible".to_string(),
                                value: Value::Boolean(false),
                            });
                            if let Ok(parent_val) = self.evaluate_expr(parent_expr) {
                                if let Value::Object(parent_ref) = &parent_val {
                                    parent_ref.borrow_mut().fields.remove(&key.to_lowercase());
                                }
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    "containskey" => {
                        // Controls.ContainsKey(key) — check if a control with that name exists
                        if !args.is_empty() {
                            let key = self.evaluate_expr(&args[0])?.as_string();
                            let controls_val = self.evaluate_expr(&Expression::MemberAccess(
                                parent_expr.clone(), member.clone()))?;
                            if let Value::Collection(coll) = &controls_val {
                                let found = coll.borrow().items.iter().any(|item| {
                                    if let Value::Object(o) = item {
                                        o.borrow().fields.get("name")
                                            .map(|v| v.as_string().eq_ignore_ascii_case(&key))
                                            .unwrap_or(false)
                                    } else { false }
                                });
                                return Ok(Value::Boolean(found));
                            }
                        }
                        return Ok(Value::Boolean(false));
                    }
                    "find" => {
                        // Controls.Find(key, searchAllChildren) — find controls by name
                        if !args.is_empty() {
                            let key = self.evaluate_expr(&args[0])?.as_string();
                            let _search_all = if args.len() > 1 {
                                self.evaluate_expr(&args[1])?.as_bool().unwrap_or(false)
                            } else { false };
                            let controls_val = self.evaluate_expr(&Expression::MemberAccess(
                                parent_expr.clone(), member.clone()))?;
                            let mut result = Vec::new();
                            if let Value::Collection(coll) = &controls_val {
                                for item in &coll.borrow().items {
                                    if let Value::Object(o) = item {
                                        let matches = o.borrow().fields.get("name")
                                            .map(|v| v.as_string().eq_ignore_ascii_case(&key))
                                            .unwrap_or(false);
                                        if matches {
                                            result.push(item.clone());
                                        }
                                    }
                                }
                            }
                            return Ok(Value::Array(result));
                        }
                        return Ok(Value::Array(Vec::new()));
                    }
                    "indexofkey" => {
                        // Controls.IndexOfKey(key) — find index of control by name
                        if !args.is_empty() {
                            let key = self.evaluate_expr(&args[0])?.as_string();
                            let controls_val = self.evaluate_expr(&Expression::MemberAccess(
                                parent_expr.clone(), member.clone()))?;
                            if let Value::Collection(coll) = &controls_val {
                                for (i, item) in coll.borrow().items.iter().enumerate() {
                                    if let Value::Object(o) = item {
                                        let matches = o.borrow().fields.get("name")
                                            .map(|v| v.as_string().eq_ignore_ascii_case(&key))
                                            .unwrap_or(false);
                                        if matches {
                                            return Ok(Value::Integer(i as i32));
                                        }
                                    }
                                }
                            }
                        }
                        return Ok(Value::Integer(-1));
                    }
                    "contains" => {
                        // Controls.Contains(ctrl) — check if collection contains this control ref
                        if !args.is_empty() {
                            let ctrl_val = self.evaluate_expr(&args[0])?;
                            let controls_val = self.evaluate_expr(&Expression::MemberAccess(
                                parent_expr.clone(), member.clone()))?;
                            if let Value::Collection(coll) = &controls_val {
                                if let Value::Object(target) = &ctrl_val {
                                    let found = coll.borrow().items.iter().any(|item| {
                                        if let Value::Object(o) = item {
                                            std::rc::Rc::ptr_eq(o, target)
                                        } else { false }
                                    });
                                    return Ok(Value::Boolean(found));
                                }
                            }
                        }
                        return Ok(Value::Boolean(false));
                    }
                    "count" => {
                        let controls_val = self.evaluate_expr(&Expression::MemberAccess(
                            parent_expr.clone(), member.clone()))?;
                        if let Value::Collection(coll) = &controls_val {
                            return Ok(Value::Integer(coll.borrow().items.len() as i32));
                        }
                        return Ok(Value::Integer(0));
                    }
                    _ => {
                        // Fall through for any other methods — let generic Collection dispatch handle
                    }
                }
            }
        }

        // Static dispatch for MemberAccess chains representing static classes
        // (e.g. System.Drawing.ColorTranslator.FromHtml(...) where we can't
        //  evaluate "System" as a variable)
        {
            fn expr_to_static_path(e: &Expression) -> Option<String> {
                match e {
                    Expression::Variable(id) => Some(id.as_str().to_lowercase()),
                    Expression::MemberAccess(base, member) => {
                        expr_to_static_path(base).map(|b| format!("{}.{}", b, member.as_str().to_lowercase()))
                    }
                    _ => None,
                }
            }
            if let Some(class_path) = expr_to_static_path(obj) {
                match class_path.as_str() {
                    "colortranslator" | "system.drawing.colortranslator" => {
                        let arg_values: Vec<Value> = args.iter().map(|a| self.evaluate_expr(a)).collect::<Result<_,_>>()?;
                        match method_name.as_str() {
                            "fromhtml" => {
                                // Return the HTML color string as-is (the interpreter treats colors as strings)
                                let color_str = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                return Ok(Value::String(color_str));
                            }
                            "tohtml" => {
                                return Ok(arg_values.get(0).map(|v| v.as_string()).map(Value::String).unwrap_or(Value::String(String::new())));
                            }
                            _ => return Ok(Value::Nothing),
                        }
                    }
                    _ => {}
                }
            }
        }

        // Evaluate object to check if it's a Collection or Dialog
        let eval_result = self.evaluate_expr(obj);
        // if method_name == "add" { ... }
        if let Ok(ref obj_val) = eval_result {
            // Universal value methods (works on any type: Integer, String, Double, Boolean, etc.)
            match method_name.as_str() {
                "tostring" => {
                    // For typed objects, check special ToString implementations
                    if let Value::Object(obj_ref) = &obj_val {
                        let tn = obj_ref.borrow().fields.get("__type").and_then(|v| {
                            if let Value::String(s) = v { Some(s.clone()) } else { None }
                        }).unwrap_or_default();
                        if tn == "Guid" {
                            let val = obj_ref.borrow().fields.get("__value").cloned().unwrap_or(Value::String(String::new()));
                            return Ok(val);
                        }
                        if tn == "Stopwatch" {
                            let elapsed = obj_ref.borrow().fields.get("elapsedmilliseconds").cloned().unwrap_or(Value::Long(0));
                            return Ok(Value::String(format!("{}ms", elapsed.as_string())));
                        }
                        if tn == "StringBuilder" {
                            let arg_values: Vec<Value> = args.iter()
                                .map(|arg| self.evaluate_expr(arg))
                                .collect::<Result<Vec<_>, _>>()?;
                            return stringbuilder_method_fn("tostring", &obj_val, &arg_values);
                        }
                        }
                    // XML objects: delegate to xml module for proper serialization
                    if crate::builtins::xml::is_xml_object(&obj_val) {
                        return crate::builtins::xml::xml_method_call(&obj_val, "tostring", &[]);
                    }
                    if let Value::Date(ole) = &obj_val {
                         let arg_values: Vec<Value> = args.iter().map(|a| self.evaluate_expr(a)).collect::<Result<_,_>>()?;
                         let fmt = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                         return Ok(Value::String(format_ole_date(*ole, &fmt)));
                    }
                    return Ok(Value::String(obj_val.as_string()));
                }
                "gethashcode" => return Ok(Value::Integer(obj_val.as_string().len() as i32)),
                "gettype" => {
                    let (type_name, full_name) = match &obj_val {
                        Value::Integer(_) => ("Int32", "System.Int32"),
                        Value::Long(_) => ("Int64", "System.Int64"),
                        Value::Single(_) => ("Single", "System.Single"),
                        Value::Double(_) => ("Double", "System.Double"),
                        Value::String(_) => ("String", "System.String"),
                        Value::Boolean(_) => ("Boolean", "System.Boolean"),
                        Value::Byte(_) => ("Byte", "System.Byte"),
                        Value::Char(_) => ("Char", "System.Char"),
                        Value::Date(_) => ("DateTime", "System.DateTime"),
                        Value::Array(_) => ("Array", "System.Array"),
                        Value::Nothing => ("Object", "System.Object"),
                        Value::Object(rc) => {
                            let cn = rc.borrow().class_name.clone();
                            let full = format!("System.{}", cn);
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("name".to_string(), Value::String(cn.clone()));
                            fields.insert("fullname".to_string(), Value::String(full));
                            fields.insert("namespace".to_string(), Value::String("System".to_string()));
                            fields.insert("__type".to_string(), Value::String("Type".to_string()));
                            let type_obj = crate::value::ObjectData {
                                class_name: "Type".to_string(),
                                fields,
                            };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(type_obj))));
                        }
                        _ => ("Object", "System.Object"),
                    };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("name".to_string(), Value::String(type_name.to_string()));
                    fields.insert("fullname".to_string(), Value::String(full_name.to_string()));
                    fields.insert("namespace".to_string(), Value::String("System".to_string()));
                    fields.insert("__type".to_string(), Value::String("Type".to_string()));
                    let type_obj = crate::value::ObjectData {
                        class_name: "Type".to_string(),
                        fields,
                    };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(type_obj))));
                }
                "equals" => {
                    // Delegate to StringBuilder-specific Equals
                    if let Value::Object(obj_ref) = &obj_val {
                        let tn = obj_ref.borrow().fields.get("__type").and_then(|v| {
                            if let Value::String(s) = v { Some(s.clone()) } else { None }
                        }).unwrap_or_default();
                        if tn == "StringBuilder" {
                            let arg_values: Vec<Value> = args.iter()
                                .map(|arg| self.evaluate_expr(arg))
                                .collect::<Result<Vec<_>, _>>()?;
                            return stringbuilder_method_fn("equals", &obj_val, &arg_values);
                        }
                    }
                    let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                        .map(|arg| self.evaluate_expr(arg))
                        .collect();
                    let arg_values = arg_values?;
                    let other = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                    return Ok(Value::Boolean(obj_val.as_string() == other.as_string()));
                }
                "compareto" => {
                    let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                        .map(|arg| self.evaluate_expr(arg))
                        .collect();
                    let arg_values = arg_values?;
                    let other = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                    let cmp = obj_val.as_string().cmp(&other.as_string());
                    return Ok(Value::Integer(match cmp {
                        std::cmp::Ordering::Less => -1,
                        std::cmp::Ordering::Equal => 0,
                        std::cmp::Ordering::Greater => 1,
                    }));
                }
                _ => {}
            }

            // Handle XML object methods: Element(), Elements(), Attribute(), Add(), etc.
            if crate::builtins::xml::is_xml_object(&obj_val) {
                let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                    .map(|arg| self.evaluate_expr(arg))
                    .collect();
                let arg_values = arg_values?;
                return crate::builtins::xml::xml_method_call(&obj_val, &method_name, &arg_values);
            }

            // Handle ResourceManager methods: GetString(key), GetObject(key)
            if let Value::Object(obj_ref) = &obj_val {
                let cn = obj_ref.borrow().class_name.clone();
                if cn == "ResourceManager" {
                    let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                        .map(|arg| self.evaluate_expr(arg))
                        .collect();
                    let arg_values = arg_values?;
                    match method_name.as_str() {
                        "getstring" => {
                            let key = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                            let lookup = format!("__res_{}", key.to_lowercase());
                            let val = obj_ref.borrow().fields.get(&lookup).cloned()
                                .unwrap_or(Value::Nothing);
                            return Ok(val);
                        }
                        "getobject" => {
                            // GetObject returns the value from the parent My.Resources (could be file object)
                            let key = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                            // Look up in our own fields first
                            let lookup = format!("__res_{}", key.to_lowercase());
                            let val = obj_ref.borrow().fields.get(&lookup).cloned()
                                .unwrap_or(Value::Nothing);
                            return Ok(val);
                        }
                        _ => {}
                    }
                }
            }

            // Handle DateTime instance methods (Value::Date(f64) - OLE Automation Date)
            if let Value::Date(ole_val) = &obj_val {
                let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                    .map(|arg| self.evaluate_expr(arg))
                    .collect();
                let arg_values = arg_values?;
                let ndt = ole_to_dt(*ole_val);
                use chrono::{Datelike, Timelike, NaiveDate};
                match method_name.as_str() {
                    // Properties (accessed as methods with no args)
                    "year" => return Ok(Value::Integer(ndt.year())),
                    "month" => return Ok(Value::Integer(ndt.month() as i32)),
                    "day" => return Ok(Value::Integer(ndt.day() as i32)),
                    "hour" => return Ok(Value::Integer(ndt.hour() as i32)),
                    "minute" => return Ok(Value::Integer(ndt.minute() as i32)),
                    "second" => return Ok(Value::Integer(ndt.second() as i32)),
                    "millisecond" => return Ok(Value::Integer((ndt.nanosecond() / 1_000_000) as i32)),
                    "dayofweek" => return Ok(Value::Integer(ndt.weekday().num_days_from_sunday() as i32)),
                    "dayofyear" => return Ok(Value::Integer(ndt.ordinal() as i32)),
                    "date" => {
                        let d = ndt.date().and_hms_opt(0, 0, 0).unwrap();
                        return Ok(Value::Date(date_to_ole(d)));
                    }
                    "timeofday" => {
                        let total_seconds = (ndt.hour() as f64) * 3600.0 + (ndt.minute() as f64) * 60.0 + (ndt.second() as f64);
                        let obj_data = crate::value::ObjectData {
                            class_name: "TimeSpan".to_string(),
                            fields: {
                                let mut f = std::collections::HashMap::new();
                                f.insert("days".to_string(), Value::Integer(0));
                                f.insert("hours".to_string(), Value::Integer(ndt.hour() as i32));
                                f.insert("minutes".to_string(), Value::Integer(ndt.minute() as i32));
                                f.insert("seconds".to_string(), Value::Integer(ndt.second() as i32));
                                f.insert("milliseconds".to_string(), Value::Integer(0));
                                f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                                f.insert("totalmilliseconds".to_string(), Value::Double(total_seconds * 1000.0));
                                f.insert("totalminutes".to_string(), Value::Double(total_seconds / 60.0));
                                f.insert("totalhours".to_string(), Value::Double(total_seconds / 3600.0));
                                f.insert("totaldays".to_string(), Value::Double(total_seconds / 86400.0));
                                f
                            },
                        };
                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
                    }
                    "ticks" => {
                        let epoch = NaiveDate::from_ymd_opt(1, 1, 1).unwrap().and_hms_opt(0, 0, 0).unwrap();
                        let diff = ndt.signed_duration_since(epoch);
                        let ticks = diff.num_milliseconds() * 10_000;
                        return Ok(Value::Long(ticks));
                    }
                    // Add methods — manipulate OLE f64 directly where possible
                    "adddays" => {
                        let n = arg_values.get(0).map(|v| v.as_double().unwrap_or(0.0)).unwrap_or(0.0);
                        return Ok(Value::Date(*ole_val + n));
                    }
                    "addhours" => {
                        let n = arg_values.get(0).map(|v| v.as_double().unwrap_or(0.0)).unwrap_or(0.0);
                        return Ok(Value::Date(*ole_val + n / 24.0));
                    }
                    "addminutes" => {
                        let n = arg_values.get(0).map(|v| v.as_double().unwrap_or(0.0)).unwrap_or(0.0);
                        return Ok(Value::Date(*ole_val + n / 1440.0));
                    }
                    "addseconds" => {
                        let n = arg_values.get(0).map(|v| v.as_double().unwrap_or(0.0)).unwrap_or(0.0);
                        return Ok(Value::Date(*ole_val + n / 86400.0));
                    }
                    "addmilliseconds" => {
                        let n = arg_values.get(0).map(|v| v.as_double().unwrap_or(0.0)).unwrap_or(0.0);
                        return Ok(Value::Date(*ole_val + n / 86400000.0));
                    }
                    "addmonths" => {
                        let n = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0);
                        let total_months = (ndt.year() * 12 + ndt.month() as i32 - 1) + n;
                        let new_year = total_months / 12;
                        let new_month = (total_months % 12 + 1) as u32;
                        let max_day = if new_month == 12 {
                            31
                        } else {
                            match (NaiveDate::from_ymd_opt(new_year, new_month, 1), NaiveDate::from_ymd_opt(if new_month == 12 { new_year + 1 } else { new_year }, if new_month == 12 { 1 } else { new_month + 1 }, 1)) {
                                (Some(a), Some(b)) => (b - a).num_days() as u32,
                                _ => 30,
                            }
                        };
                        let new_day = std::cmp::min(ndt.day(), max_day);
                        let new_dt = NaiveDate::from_ymd_opt(new_year, new_month, new_day)
                            .unwrap_or(ndt.date())
                            .and_hms_opt(ndt.hour(), ndt.minute(), ndt.second())
                            .unwrap_or(ndt);
                        return Ok(Value::Date(date_to_ole(new_dt)));
                    }
                    "addyears" => {
                        let n = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0);
                        let new_year = ndt.year() + n;
                        let max_day = if ndt.month() == 2 && ndt.day() == 29 {
                            if NaiveDate::from_ymd_opt(new_year, 2, 29).is_some() { 29 } else { 28 }
                        } else {
                            ndt.day()
                        };
                        let new_dt = NaiveDate::from_ymd_opt(new_year, ndt.month(), max_day)
                            .unwrap_or(ndt.date())
                            .and_hms_opt(ndt.hour(), ndt.minute(), ndt.second())
                            .unwrap_or(ndt);
                        return Ok(Value::Date(date_to_ole(new_dt)));
                    }
                    // Subtract — returns TimeSpan
                    "subtract" => {
                        if let Some(Value::Date(other_ole)) = arg_values.get(0) {
                            let diff_days = *ole_val - *other_ole;
                            let total_seconds = diff_days * 86400.0;
                            let obj_data = crate::value::ObjectData {
                                class_name: "TimeSpan".to_string(),
                                fields: {
                                    let mut f = std::collections::HashMap::new();
                                    f.insert("days".to_string(), Value::Integer(diff_days.trunc() as i32));
                                    f.insert("hours".to_string(), Value::Integer(((total_seconds % 86400.0) / 3600.0) as i32));
                                    f.insert("minutes".to_string(), Value::Integer(((total_seconds % 3600.0) / 60.0) as i32));
                                    f.insert("seconds".to_string(), Value::Integer((total_seconds % 60.0) as i32));
                                    f.insert("milliseconds".to_string(), Value::Integer(((total_seconds * 1000.0) % 1000.0) as i32));
                                    f.insert("totaldays".to_string(), Value::Double(diff_days));
                                    f.insert("totalhours".to_string(), Value::Double(total_seconds / 3600.0));
                                    f.insert("totalminutes".to_string(), Value::Double(total_seconds / 60.0));
                                    f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                                    f.insert("totalmilliseconds".to_string(), Value::Double(total_seconds * 1000.0));
                                    f
                                },
                            };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
                        }
                        return Err(RuntimeError::Custom("DateTime.Subtract requires a DateTime argument".to_string()));
                    }

                    "toshortdatestring" => return Ok(Value::String(ole_to_dt(*ole_val).format("%m/%d/%Y").to_string())),
                    "tolongdatestring" => return Ok(Value::String(ole_to_dt(*ole_val).format("%A, %B %d, %Y").to_string())),
                    "toshorttimestring" => return Ok(Value::String(ole_to_dt(*ole_val).format("%H:%M").to_string())),
                    "tolongtimestring" => return Ok(Value::String(ole_to_dt(*ole_val).format("%H:%M:%S").to_string())),
                    "tofiletime" => {
                        // File time = 100-nanosecond intervals since 1601-01-01
                        let epoch_1601 = NaiveDate::from_ymd_opt(1601, 1, 1).unwrap().and_hms_opt(0, 0, 0).unwrap();
                        let diff = ndt.signed_duration_since(epoch_1601);
                        return Ok(Value::Long(diff.num_milliseconds() * 10_000));
                    }
                    _ => {} // Fall through to other dispatch
                }
            }

            // ===== .NET String instance methods =====
            if let Value::String(s_val) = obj_val {
                let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                    .map(|arg| self.evaluate_expr(arg))
                    .collect();
                let arg_values = arg_values?;
                match method_name.as_str() {
                    "tolower" | "tolowerinvariant" => return Ok(Value::String(s_val.to_lowercase())),
                    "toupper" | "toupperinvariant" => return Ok(Value::String(s_val.to_uppercase())),
                    "trim" => {
                        if let Some(chars_arg) = arg_values.get(0) {
                            let trim_chars: Vec<char> = chars_arg.as_string().chars().collect();
                            return Ok(Value::String(s_val.trim_matches(|c| trim_chars.contains(&c)).to_string()));
                        }
                        return Ok(Value::String(s_val.trim().to_string()));
                    }
                    "trimstart" | "trimend" => {
                        let trim_chars: Vec<char> = arg_values.get(0)
                            .map(|v| v.as_string().chars().collect())
                            .unwrap_or_else(|| vec![' ', '\t', '\n', '\r']);
                        if method_name == "trimstart" {
                            return Ok(Value::String(s_val.trim_start_matches(|c| trim_chars.contains(&c)).to_string()));
                        } else {
                            return Ok(Value::String(s_val.trim_end_matches(|c| trim_chars.contains(&c)).to_string()));
                        }
                    }
                    "contains" => {
                        let needle = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                        return Ok(Value::Boolean(s_val.contains(&needle)));
                    }
                    "startswith" => {
                        let prefix = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                        // Check for StringComparison.OrdinalIgnoreCase (=5)
                        let ignore_case = arg_values.get(1).and_then(|v| v.as_integer().ok()).map(|i| i == 5).unwrap_or(false);
                        if ignore_case {
                            return Ok(Value::Boolean(s_val.to_lowercase().starts_with(&prefix.to_lowercase())));
                        }
                        return Ok(Value::Boolean(s_val.starts_with(&prefix)));
                    }
                    "endswith" => {
                        let suffix = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                        let ignore_case = arg_values.get(1).and_then(|v| v.as_integer().ok()).map(|i| i == 5).unwrap_or(false);
                        if ignore_case {
                            return Ok(Value::Boolean(s_val.to_lowercase().ends_with(&suffix.to_lowercase())));
                        }
                        return Ok(Value::Boolean(s_val.ends_with(&suffix)));
                    }
                    "indexof" => {
                        let needle = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                        let start = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                        let hay = if start > 0 && start < s_val.len() { &s_val[start..] } else { s_val.as_str() };
                        match hay.find(&needle) {
                            Some(pos) => return Ok(Value::Integer((pos + start) as i32)),
                            None => return Ok(Value::Integer(-1)),
                        }
                    }
                    "lastindexof" => {
                        let needle = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                        match s_val.rfind(&needle) {
                            Some(pos) => return Ok(Value::Integer(pos as i32)),
                            None => return Ok(Value::Integer(-1)),
                        }
                    }
                    "substring" => {
                        let start = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                        let len = arg_values.get(1).and_then(|v| v.as_integer().ok());
                        let chars: Vec<char> = s_val.chars().collect();
                        if start >= chars.len() {
                            return Ok(Value::String(String::new()));
                        }
                        let result: String = if let Some(l) = len {
                            chars[start..std::cmp::min(start + l as usize, chars.len())].iter().collect()
                        } else {
                            chars[start..].iter().collect()
                        };
                        return Ok(Value::String(result));
                    }
                    "replace" => {
                        let old = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                        let new = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                        return Ok(Value::String(s_val.replace(&old, &new)));
                    }
                    "remove" => {
                        let start = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                        let count = arg_values.get(1).and_then(|v| v.as_integer().ok());
                        let chars: Vec<char> = s_val.chars().collect();
                        let result: String = if let Some(c) = count {
                            let c = c as usize;
                            chars[..start].iter().chain(chars[std::cmp::min(start + c, chars.len())..].iter()).collect()
                        } else {
                            chars[..start].iter().collect()
                        };
                        return Ok(Value::String(result));
                    }

                    "insert" => {
                        let idx = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                        let ins = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                        let mut result = s_val.clone();
                        if idx <= result.len() {
                            result.insert_str(idx, &ins);
                        }
                        return Ok(Value::String(result));
                    }
                    "split" => {
                        // .Split(Char) / .Split(String()) / .Split({delims}, StringSplitOptions)
                        let delim = arg_values.get(0).cloned().unwrap_or(Value::String(" ".to_string()));
                        // Check for StringSplitOptions (second arg): 0=None, 1=RemoveEmptyEntries
                        let remove_empty = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0) == 1;
                        let parts: Vec<&str> = match &delim {
                            Value::Array(arr) => {
                                // Array of string delimiters
                                let delims: Vec<String> = arr.iter().map(|v| v.as_string()).collect();
                                if delims.len() == 1 {
                                    s_val.split(&delims[0]).collect()
                                } else {
                                    // Multi-delimiter split: use first as primary, then split results by others
                                    // For multi-string delimiters, we need owned strings
                                    let mut owned_parts: Vec<String> = vec![s_val.clone()];
                                    for d in &delims {
                                        let mut new_parts = Vec::new();
                                        for part in &owned_parts {
                                            for sub in part.split(d.as_str()) {
                                                new_parts.push(sub.to_string());
                                            }
                                        }
                                        owned_parts = new_parts;
                                    }
                                    let result_vals: Vec<Value> = owned_parts.into_iter()
                                        .filter(|s| !remove_empty || !s.is_empty())
                                        .map(|s| Value::String(s))
                                        .collect();
                                    return Ok(Value::Array(result_vals));
                                }
                            }
                            Value::Char(c) => s_val.split(*c).collect(),
                            _ => {
                                let d = delim.as_string();
                                if d.len() == 1 {
                                    s_val.split(d.chars().next().unwrap()).collect()
                                } else {
                                    s_val.split(&d).collect()
                                }
                            }
                        };
                        let result: Vec<Value> = parts.into_iter()
                            .filter(|s| !remove_empty || !s.is_empty())
                            .map(|s| Value::String(s.to_string()))
                            .collect();
                        return Ok(Value::Array(result));
                    }
                    "padleft" => {
                        let total_width = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                        let pad_char = arg_values.get(1).map(|v| {
                            let s = v.as_string();
                            s.chars().next().unwrap_or(' ')
                        }).unwrap_or(' ');
                        if s_val.len() >= total_width {
                            return Ok(Value::String(s_val.clone()));
                        }
                        let padding: String = std::iter::repeat(pad_char).take(total_width - s_val.len()).collect();
                        return Ok(Value::String(format!("{}{}", padding, s_val)));
                    }
                    "padright" => {
                        let total_width = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                        let pad_char = arg_values.get(1).map(|v| {
                            let s = v.as_string();
                            s.chars().next().unwrap_or(' ')
                        }).unwrap_or(' ');
                        if s_val.len() >= total_width {
                            return Ok(Value::String(s_val.clone()));
                        }
                        let padding: String = std::iter::repeat(pad_char).take(total_width - s_val.len()).collect();
                        return Ok(Value::String(format!("{}{}", s_val, padding)));
                    }
                    "chars" => {
                        // .Chars(index) — returns character at index
                        let idx = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                        let c = s_val.chars().nth(idx).unwrap_or('\0');
                        return Ok(Value::Char(c));
                    }
                    "tochararray" => {
                        let arr: Vec<Value> = s_val.chars().map(|c| Value::Char(c)).collect();
                        return Ok(Value::Array(arr));
                    }
                    "isnullorempty" => {
                        return Ok(Value::Boolean(s_val.is_empty()));
                    }
                    "tostring" => {
                        return Ok(Value::String(s_val.clone()));
                    }
                    _ => {} // Fall through to other dispatch
                }
            }

            // ===== Array instance methods =====
            if let Value::Array(arr_val) = obj_val {
                let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                    .map(|arg| self.evaluate_expr(arg))
                    .collect();
                let arg_values = arg_values?;
                match method_name.as_str() {
                    "contains" => {
                        let needle = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                        let found = arr_val.iter().any(|v| v.as_string() == needle.as_string());
                        return Ok(Value::Boolean(found));
                    }
                    "tostring" => {
                        let joined = arr_val.iter().map(|v| v.as_string()).collect::<Vec<_>>().join(", ");
                        return Ok(Value::String(format!("[{}]", joined)));
                    }
                    _ => {} // Fall through
                }
            }

            // Handle StringBuilder methods
            if let Value::Object(obj_ref) = &obj_val {
                let type_name = obj_ref.borrow().fields.get("__type").and_then(|v| {
                    if let Value::String(s) = v { Some(s.clone()) } else { None }
                }).unwrap_or_default();
                if !type_name.is_empty() {
                    if type_name == "StringBuilder" {
                        // Special handling for CopyTo: write back the modified array to the destination variable
                        if method_name == "copyto" && args.len() >= 4 {
                            let arg_values: Vec<Value> = args.iter()
                                .map(|arg| self.evaluate_expr(arg))
                                .collect::<Result<Vec<_>, _>>()?;
                            let result = stringbuilder_method_fn(&method_name, &obj_val, &arg_values)?;
                            // Write the result array back to the destination variable
                            if let Value::Array(_) = &result {
                                if let Expression::Variable(dest_id) = &args[1] {
                                    self.env.set(dest_id.as_str(), result.clone())?;
                                }
                            }
                            return Ok(result);
                        }
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                            .map(|arg| self.evaluate_expr(arg))
                            .collect();
                        return stringbuilder_method_fn(&method_name, &obj_val, &arg_values?);
                    }
                    if type_name == "WebClient" {
                        // Collect headers from the object for curl -H flags
                        let header_args: Vec<String> = {
                            if let Some(Value::Dictionary(hd)) = obj_ref.borrow().fields.get("headers") {
                                let hd_b = hd.borrow();
                                let keys = hd_b.keys();
                                let vals = hd_b.values();
                                keys.iter().zip(vals.iter()).map(|(k, v)| {
                                    format!("{}: {}", k.as_string(), v.as_string())
                                }).collect()
                            } else {
                                vec![]
                            }
                        };
                        // Check for credentials
                        let cred_args: Vec<String> = {
                            if let Some(Value::Object(cred_ref)) = obj_ref.borrow().fields.get("credentials") {
                                let cb = cred_ref.borrow();
                                if let (Some(Value::String(u)), Some(Value::String(p))) = (cb.fields.get("username"), cb.fields.get("password")) {
                                    vec!["-u".to_string(), format!("{}:{}", u, p)]
                                } else { vec![] }
                            } else { vec![] }
                        };
                        let build_curl = |extra: &[&str], url: &str| -> std::process::Command {
                            let mut cmd = std::process::Command::new("curl");
                            cmd.arg("-s").arg("-L");
                            for h in &header_args {
                                cmd.arg("-H").arg(h);
                            }
                            for c in &cred_args {
                                cmd.arg(c);
                            }
                            for e in extra {
                                cmd.arg(*e);
                            }
                            cmd.arg(url);
                            cmd
                        };
                        match method_name.as_str() {
                            "downloadstring" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let output = build_curl(&[], &url).output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                if !output.status.success() {
                                    let stderr = String::from_utf8_lossy(&output.stderr);
                                    return Err(RuntimeError::Custom(format!("WebClient.DownloadString failed: {}", stderr)));
                                }
                                return Ok(Value::String(String::from_utf8_lossy(&output.stdout).to_string()));
                            }
                            "downloaddata" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let output = build_curl(&[], &url).output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                if !output.status.success() {
                                    let stderr = String::from_utf8_lossy(&output.stderr);
                                    return Err(RuntimeError::Custom(format!("WebClient.DownloadData failed: {}", stderr)));
                                }
                                let bytes: Vec<Value> = output.stdout.iter().map(|b| Value::Byte(*b)).collect();
                                return Ok(Value::Array(bytes));
                            }
                            "downloadfile" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let path = self.evaluate_expr(&args[1])?.as_string();
                                let output = build_curl(&["-o", &path], &url).output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                if !output.status.success() {
                                    let stderr = String::from_utf8_lossy(&output.stderr);
                                    return Err(RuntimeError::Custom(format!("WebClient.DownloadFile failed: {}", stderr)));
                                }
                                return Ok(Value::Nothing);
                            }
                            "uploadstring" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let data = self.evaluate_expr(&args[1])?.as_string();
                                let output = build_curl(&["-X", "POST", "-d", &data], &url).output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                if !output.status.success() {
                                    let stderr = String::from_utf8_lossy(&output.stderr);
                                    return Err(RuntimeError::Custom(format!("WebClient.UploadString failed: {}", stderr)));
                                }
                                return Ok(Value::String(String::from_utf8_lossy(&output.stdout).to_string()));
                            }
                            "uploaddata" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let data = self.evaluate_expr(&args[1])?;
                                // Convert byte array to a temp approach: write via stdin
                                let bytes: Vec<u8> = if let Value::Array(arr) = &data {
                                    arr.iter().map(|v| match v { Value::Byte(b) => *b, _ => 0 }).collect()
                                } else {
                                    data.as_string().into_bytes()
                                };
                                use std::io::Write;
                                let mut child = std::process::Command::new("curl")
                                    .args(&["-s", "-L", "-X", "POST", "--data-binary", "@-"])
                                    .args(header_args.iter().flat_map(|h| vec!["-H", h.as_str()]))
                                    .args(&cred_args)
                                    .arg(&url)
                                    .stdin(std::process::Stdio::piped())
                                    .stdout(std::process::Stdio::piped())
                                    .stderr(std::process::Stdio::piped())
                                    .spawn()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                if let Some(mut stdin) = child.stdin.take() {
                                    let _ = stdin.write_all(&bytes);
                                }
                                let output = child.wait_with_output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                let resp_bytes: Vec<Value> = output.stdout.iter().map(|b| Value::Byte(*b)).collect();
                                return Ok(Value::Array(resp_bytes));
                            }
                            "dispose" => return Ok(Value::Nothing),
                            _ => {}
                        }
                    }
                    if type_name == "HttpClient" {
                        // Collect DefaultRequestHeaders
                        let header_args: Vec<String> = {
                            if let Some(Value::Dictionary(hd)) = obj_ref.borrow().fields.get("defaultrequestheaders") {
                                let hd_b = hd.borrow();
                                let keys = hd_b.keys();
                                let vals = hd_b.values();
                                keys.iter().zip(vals.iter()).map(|(k, v)| {
                                    format!("{}: {}", k.as_string(), v.as_string())
                                }).collect()
                            } else {
                                vec![]
                            }
                        };
                        let build_http_curl = |method: &str, url: &str, body: Option<&str>, content_type: Option<&str>| -> std::process::Command {
                            let mut cmd = std::process::Command::new("curl");
                            cmd.arg("-s").arg("-L").arg("-w").arg("\n%{http_code}");
                            cmd.arg("-X").arg(method);
                            for h in &header_args {
                                cmd.arg("-H").arg(h);
                            }
                            if let Some(ct) = content_type {
                                cmd.arg("-H").arg(&format!("Content-Type: {}", ct));
                            }
                            if let Some(b) = body {
                                cmd.arg("-d").arg(b);
                            }
                            cmd.arg(url);
                            cmd
                        };
                        let make_response = |output: std::process::Output| -> Result<Value, RuntimeError> {
                            let full = String::from_utf8_lossy(&output.stdout).to_string();
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("__type".to_string(), Value::String("HttpResponseMessage".to_string()));
                            if let Some(pos) = full.rfind('\n') {
                                let body = full[..pos].to_string();
                                let code = full[pos+1..].trim().to_string();
                                fields.insert("content".to_string(), Value::String(body));
                                fields.insert("statuscode".to_string(), Value::String(code.clone()));
                                fields.insert("issuccessstatuscode".to_string(), Value::Boolean(code.starts_with('2')));
                            } else {
                                fields.insert("content".to_string(), Value::String(full));
                                fields.insert("statuscode".to_string(), Value::String("200".to_string()));
                                fields.insert("issuccessstatuscode".to_string(), Value::Boolean(true));
                            }
                            let obj = crate::value::ObjectData { class_name: "HttpResponseMessage".to_string(), fields };
                            Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))))
                        };
                        match method_name.as_str() {
                            "getstringasync" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let output = build_http_curl("GET", &url, None, None).output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                // getstringasync returns the body directly, not a response object
                                let full = String::from_utf8_lossy(&output.stdout).to_string();
                                // Strip the status code line we appended
                                let body = if let Some(pos) = full.rfind('\n') {
                                    full[..pos].to_string()
                                } else { full };
                                return Ok(Value::String(body));
                            }
                            "getasync" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let output = build_http_curl("GET", &url, None, None).output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                return make_response(output);
                            }
                            "postasync" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let content_val = self.evaluate_expr(&args[1])?;
                                let (body_str, ct) = if let Value::Object(cref) = &content_val {
                                    let cb = cref.borrow();
                                    let b = cb.fields.get("body").map(|v| v.as_string()).unwrap_or_default();
                                    let t = cb.fields.get("mediatype").map(|v| v.as_string()).unwrap_or("application/json".to_string());
                                    (b, t)
                                } else {
                                    (content_val.as_string(), "text/plain".to_string())
                                };
                                let output = build_http_curl("POST", &url, Some(&body_str), Some(&ct)).output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                return make_response(output);
                            }
                            "putasync" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let content_val = self.evaluate_expr(&args[1])?;
                                let (body_str, ct) = if let Value::Object(cref) = &content_val {
                                    let cb = cref.borrow();
                                    let b = cb.fields.get("body").map(|v| v.as_string()).unwrap_or_default();
                                    let t = cb.fields.get("mediatype").map(|v| v.as_string()).unwrap_or("application/json".to_string());
                                    (b, t)
                                } else {
                                    (content_val.as_string(), "text/plain".to_string())
                                };
                                let output = build_http_curl("PUT", &url, Some(&body_str), Some(&ct)).output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                return make_response(output);
                            }
                            "deleteasync" => {
                                let url = self.evaluate_expr(&args[0])?.as_string();
                                let output = build_http_curl("DELETE", &url, None, None).output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                return make_response(output);
                            }
                            "dispose" => return Ok(Value::Nothing),
                            _ => {}
                        }
                    }
                    // HttpResponseMessage.Content.ReadAsStringAsync()
                    if type_name == "HttpResponseMessage" {
                        match method_name.as_str() {
                            "readasstringasync" | "tostring" => {
                                let content = obj_ref.borrow().fields.get("content").cloned()
                                    .unwrap_or(Value::String(String::new()));
                                return Ok(content);
                            }
                            _ => {}
                        }
                    }
                    // HttpWebRequest methods
                    if type_name == "HttpWebRequest" {
                        match method_name.as_str() {
                            "getresponse" | "getresponseasync" => {
                                let url = obj_ref.borrow().fields.get("url").cloned()
                                    .unwrap_or(Value::String(String::new())).as_string();
                                let method_str = obj_ref.borrow().fields.get("method").cloned()
                                    .unwrap_or(Value::String("GET".to_string())).as_string();
                                let mut cmd = std::process::Command::new("curl");
                                cmd.args(&["-s", "-L", "-w", "\n%{http_code}", "-X", &method_str]);
                                // Add headers if any
                                if let Some(Value::Dictionary(hd)) = obj_ref.borrow().fields.get("headers") {
                                    let hd_b = hd.borrow();
                                    let keys = hd_b.keys();
                                    let vals = hd_b.values();
                                    for (k, v) in keys.iter().zip(vals.iter()) {
                                        cmd.arg("-H").arg(&format!("{}: {}", k.as_string(), v.as_string()));
                                    }
                                }
                                // Credentials
                                if let Some(Value::Object(cred)) = obj_ref.borrow().fields.get("credentials") {
                                    let cb = cred.borrow();
                                    if let (Some(Value::String(u)), Some(Value::String(p))) = (cb.fields.get("username"), cb.fields.get("password")) {
                                        cmd.arg("-u").arg(&format!("{}:{}", u, p));
                                    }
                                }
                                cmd.arg(&url);
                                let output = cmd.output()
                                    .map_err(|e| RuntimeError::Custom(format!("curl failed: {}", e)))?;
                                let full = String::from_utf8_lossy(&output.stdout).to_string();
                                let mut fields = std::collections::HashMap::new();
                                fields.insert("__type".to_string(), Value::String("HttpWebResponse".to_string()));
                                if let Some(pos) = full.rfind('\n') {
                                    let body = full[..pos].to_string();
                                    let code = full[pos+1..].trim().to_string();
                                    fields.insert("body".to_string(), Value::String(body.clone()));
                                    fields.insert("statuscode".to_string(), Value::String(code));
                                    fields.insert("contentlength".to_string(), Value::Integer(body.len() as i32));
                                } else {
                                    fields.insert("body".to_string(), Value::String(full.clone()));
                                    fields.insert("statuscode".to_string(), Value::String("200".to_string()));
                                    fields.insert("contentlength".to_string(), Value::Integer(full.len() as i32));
                                }
                                let obj = crate::value::ObjectData { class_name: "HttpWebResponse".to_string(), fields };
                                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                            }
                            "dispose" | "abort" => return Ok(Value::Nothing),
                            _ => {}
                        }
                    }
                    // HttpWebResponse methods
                    if type_name == "HttpWebResponse" {
                        match method_name.as_str() {
                            "getresponsestream" => {
                                // Return the body as a string (simulates reading a stream)
                                let body = obj_ref.borrow().fields.get("body").cloned()
                                    .unwrap_or(Value::String(String::new()));
                                return Ok(body);
                            }
                            "close" | "dispose" => return Ok(Value::Nothing),
                            _ => {}
                        }
                    }

                    // Random instance methods
                    if type_name == "Random" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                            .map(|arg| self.evaluate_expr(arg))
                            .collect();
                        let arg_values = arg_values?;
                        match method_name.as_str() {
                            "next" => {
                                // Simple LCG random number generator
                                let seed = obj_ref.borrow().fields.get("__seed").map(|v| {
                                    if let Value::Long(l) = v { *l as u64 } else { 0u64 }
                                }).unwrap_or(0);
                                let counter = obj_ref.borrow().fields.get("__counter").map(|v| {
                                    if let Value::Long(l) = v { *l as u64 } else { 0u64 }
                                }).unwrap_or(0);
                                let state = seed.wrapping_add(counter).wrapping_mul(6364136223846793005).wrapping_add(1442695040888963407);
                                let raw = ((state >> 33) ^ state) as i64;
                                let raw = raw.unsigned_abs();
                                obj_ref.borrow_mut().fields.insert("__counter".to_string(), Value::Long((counter + 1) as i64));
                                match arg_values.len() {
                                    0 => {
                                        return Ok(Value::Integer((raw % (i32::MAX as u64)) as i32));
                                    }
                                    1 => {
                                        let max = arg_values[0].as_integer()? as u64;
                                        if max == 0 { return Ok(Value::Integer(0)); }
                                        return Ok(Value::Integer((raw % max) as i32));
                                    }
                                    _ => {
                                        let min = arg_values[0].as_integer()?;
                                        let max = arg_values[1].as_integer()?;
                                        let range = (max - min) as u64;
                                        if range == 0 { return Ok(Value::Integer(min)); }
                                        return Ok(Value::Integer(min + (raw % range) as i32));
                                    }
                                }
                            }
                            "nextdouble" => {
                                let seed = obj_ref.borrow().fields.get("__seed").map(|v| {
                                    if let Value::Long(l) = v { *l as u64 } else { 0u64 }
                                }).unwrap_or(0);
                                let counter = obj_ref.borrow().fields.get("__counter").map(|v| {
                                    if let Value::Long(l) = v { *l as u64 } else { 0u64 }
                                }).unwrap_or(0);
                                let state = seed.wrapping_add(counter).wrapping_mul(6364136223846793005).wrapping_add(1442695040888963407);
                                let raw = ((state >> 33) ^ state) as u64;
                                obj_ref.borrow_mut().fields.insert("__counter".to_string(), Value::Long((counter + 1) as i64));
                                let val = (raw as f64) / (u64::MAX as f64);
                                return Ok(Value::Double(val));
                            }
                            "nextbytes" => {
                                let count = arg_values.get(0).map(|v| v.as_string().parse::<usize>().unwrap_or(0)).unwrap_or(0);
                                let mut bytes = Vec::new();
                                let seed = obj_ref.borrow().fields.get("__seed").map(|v| {
                                    if let Value::Long(l) = v { *l as u64 } else { 0u64 }
                                }).unwrap_or(0);
                                let mut counter = obj_ref.borrow().fields.get("__counter").map(|v| {
                                    if let Value::Long(l) = v { *l as u64 } else { 0u64 }
                                }).unwrap_or(0);
                                for _ in 0..count {
                                    let state = seed.wrapping_add(counter).wrapping_mul(6364136223846793005).wrapping_add(1442695040888963407);
                                    bytes.push(Value::Byte((state >> 40) as u8));
                                    counter += 1;
                                }
                                obj_ref.borrow_mut().fields.insert("__counter".to_string(), Value::Long(counter as i64));
                                return Ok(Value::Array(bytes));
                            }
                            _ => {}
                        }
                    }

                    // Stopwatch instance methods
                    if type_name == "Stopwatch" {
                        match method_name.as_str() {
                            "start" => {
                                let now_ms = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_millis() as i64;
                                obj_ref.borrow_mut().fields.insert("__start_ms".to_string(), Value::Long(now_ms));
                                obj_ref.borrow_mut().fields.insert("isrunning".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            "stop" => {
                                let is_running = obj_ref.borrow().fields.get("isrunning").map(|v| if let Value::Boolean(b) = v { *b } else { false }).unwrap_or(false);
                                if is_running {
                                    let now_ms = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_millis() as i64;
                                    let start = obj_ref.borrow().fields.get("__start_ms").map(|v| if let Value::Long(l) = v { *l } else { 0 }).unwrap_or(0);
                                    let accumulated = obj_ref.borrow().fields.get("__accumulated_ms").map(|v| if let Value::Long(l) = v { *l } else { 0 }).unwrap_or(0);
                                    let elapsed = accumulated + (now_ms - start);
                                    obj_ref.borrow_mut().fields.insert("__accumulated_ms".to_string(), Value::Long(elapsed));
                                    obj_ref.borrow_mut().fields.insert("elapsedmilliseconds".to_string(), Value::Long(elapsed));
                                    obj_ref.borrow_mut().fields.insert("isrunning".to_string(), Value::Boolean(false));
                                }
                                return Ok(Value::Nothing);
                            }
                            "reset" => {
                                obj_ref.borrow_mut().fields.insert("__accumulated_ms".to_string(), Value::Long(0));
                                obj_ref.borrow_mut().fields.insert("__start_ms".to_string(), Value::Long(0));
                                obj_ref.borrow_mut().fields.insert("elapsedmilliseconds".to_string(), Value::Long(0));
                                obj_ref.borrow_mut().fields.insert("isrunning".to_string(), Value::Boolean(false));
                                return Ok(Value::Nothing);
                            }
                            "restart" => {
                                let now_ms = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_millis() as i64;
                                obj_ref.borrow_mut().fields.insert("__accumulated_ms".to_string(), Value::Long(0));
                                obj_ref.borrow_mut().fields.insert("__start_ms".to_string(), Value::Long(now_ms));
                                obj_ref.borrow_mut().fields.insert("elapsedmilliseconds".to_string(), Value::Long(0));
                                obj_ref.borrow_mut().fields.insert("isrunning".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            _ => {}
                        }
                    }

                    // BackgroundWorker instance methods (sync fallback — single-threaded)
                    if type_name == "BackgroundWorker" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                            .map(|arg| self.evaluate_expr(arg))
                            .collect();
                        let arg_values = arg_values?;
                        let sender_val = Value::Object(obj_ref.clone());
                        match method_name.as_str() {
                            "runworkerasync" => {
                                // Set IsBusy = True
                                obj_ref.borrow_mut().fields.insert("isbusy".to_string(), Value::Boolean(true));
                                obj_ref.borrow_mut().fields.insert("cancellationpending".to_string(), Value::Boolean(false));

                                // Get the argument passed to RunWorkerAsync (if any)
                                let worker_arg = arg_values.get(0).cloned().unwrap_or(Value::Nothing);

                                // Create DoWorkEventArgs
                                let mut dw_fields = std::collections::HashMap::new();
                                dw_fields.insert("__type".to_string(), Value::String("DoWorkEventArgs".to_string()));
                                dw_fields.insert("argument".to_string(), worker_arg);
                                dw_fields.insert("cancel".to_string(), Value::Boolean(false));
                                dw_fields.insert("result".to_string(), Value::Nothing);
                                let dw_args = Value::Object(std::rc::Rc::new(std::cell::RefCell::new(
                                    crate::value::ObjectData { class_name: "DoWorkEventArgs".to_string(), fields: dw_fields }
                                )));

                                // Fire DoWork handler
                                let dowork_handler = obj_ref.borrow().fields.get("__dowork_handler").cloned().unwrap_or(Value::Nothing);
                                if let Value::String(handler_name) = &dowork_handler {
                                    if !handler_name.is_empty() {
                                        let handler_lower = handler_name.to_lowercase();
                                        if let Some(sub) = self.subs.get(&handler_lower).cloned() {
                                            let _ = self.call_user_sub(&sub, &[sender_val.clone(), dw_args.clone()], None);
                                        }
                                    }
                                } else if let Value::Lambda { .. } = &dowork_handler {
                                    let _ = self.call_lambda(dowork_handler.clone(), &[sender_val.clone(), dw_args.clone()]);
                                }

                                // Get result from DoWorkEventArgs
                                let dw_result = if let Value::Object(ref dw_ref) = dw_args {
                                    dw_ref.borrow().fields.get("result").cloned().unwrap_or(Value::Nothing)
                                } else {
                                    Value::Nothing
                                };
                                let was_cancelled = if let Value::Object(ref dw_ref) = dw_args {
                                    dw_ref.borrow().fields.get("cancel").map(|v| v.is_truthy()).unwrap_or(false)
                                } else {
                                    false
                                };

                                // Create RunWorkerCompletedEventArgs
                                let mut rw_fields = std::collections::HashMap::new();
                                rw_fields.insert("__type".to_string(), Value::String("RunWorkerCompletedEventArgs".to_string()));
                                rw_fields.insert("result".to_string(), dw_result);
                                rw_fields.insert("cancelled".to_string(), Value::Boolean(was_cancelled));
                                rw_fields.insert("error".to_string(), Value::Nothing);
                                let rw_args = Value::Object(std::rc::Rc::new(std::cell::RefCell::new(
                                    crate::value::ObjectData { class_name: "RunWorkerCompletedEventArgs".to_string(), fields: rw_fields }
                                )));

                                // Fire RunWorkerCompleted handler
                                let completed_handler = obj_ref.borrow().fields.get("__runworkercompleted_handler").cloned().unwrap_or(Value::Nothing);
                                if let Value::String(handler_name) = &completed_handler {
                                    if !handler_name.is_empty() {
                                        let handler_lower = handler_name.to_lowercase();
                                        if let Some(sub) = self.subs.get(&handler_lower).cloned() {
                                            let _ = self.call_user_sub(&sub, &[sender_val.clone(), rw_args], None);
                                        }
                                    }
                                } else if let Value::Lambda { .. } = &completed_handler {
                                    let _ = self.call_lambda(completed_handler.clone(), &[sender_val.clone(), rw_args]);
                                }

                                // Set IsBusy = False
                                obj_ref.borrow_mut().fields.insert("isbusy".to_string(), Value::Boolean(false));
                                return Ok(Value::Nothing);
                            }
                            "reportprogress" => {
                                let percent = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0);
                                let user_state = arg_values.get(1).cloned().unwrap_or(Value::Nothing);

                                // Create ProgressChangedEventArgs
                                let mut pc_fields = std::collections::HashMap::new();
                                pc_fields.insert("__type".to_string(), Value::String("ProgressChangedEventArgs".to_string()));
                                pc_fields.insert("progresspercentage".to_string(), Value::Integer(percent));
                                pc_fields.insert("userstate".to_string(), user_state);
                                let pc_args = Value::Object(std::rc::Rc::new(std::cell::RefCell::new(
                                    crate::value::ObjectData { class_name: "ProgressChangedEventArgs".to_string(), fields: pc_fields }
                                )));

                                // Fire ProgressChanged handler
                                let progress_handler = obj_ref.borrow().fields.get("__progresschanged_handler").cloned().unwrap_or(Value::Nothing);
                                if let Value::String(handler_name) = &progress_handler {
                                    if !handler_name.is_empty() {
                                        let handler_lower = handler_name.to_lowercase();
                                        if let Some(sub) = self.subs.get(&handler_lower).cloned() {
                                            let _ = self.call_user_sub(&sub, &[sender_val.clone(), pc_args], None);
                                        }
                                    }
                                } else if let Value::Lambda { .. } = &progress_handler {
                                    let _ = self.call_lambda(progress_handler.clone(), &[sender_val.clone(), pc_args]);
                                }
                                return Ok(Value::Nothing);
                            }
                            "cancelasync" => {
                                obj_ref.borrow_mut().fields.insert("cancellationpending".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            "dispose" => {
                                return Ok(Value::Nothing);
                            }
                            _ => {}
                        }
                    }

                    // Process instance methods
                    if type_name.eq_ignore_ascii_case("Process") || type_name.eq_ignore_ascii_case("System.Diagnostics.Process") {
                        match method_name.as_str() {
                            "waitforexit" => {
                                let handle_id = obj_ref.borrow().fields.get("__handle").map(|v| v.as_string()).unwrap_or_default();
                                if !handle_id.is_empty() {
                                    let mut reg = get_registry().lock().unwrap();
                                    if let Some(mut child) = reg.processes.remove(&handle_id) {
                                        if let Ok(status) = child.wait() {
                                            let code = status.code().unwrap_or(0);
                                            obj_ref.borrow_mut().fields.insert("exitcode".to_string(), Value::Integer(code));
                                        }
                                    }
                                }
                                obj_ref.borrow_mut().fields.insert("hasexited".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            "kill" | "close" | "dispose" => {
                                let handle_id = obj_ref.borrow().fields.get("__handle").map(|v| v.as_string()).unwrap_or_default();
                                if !handle_id.is_empty() {
                                    let mut reg = get_registry().lock().unwrap();
                                    if let Some(mut child) = reg.processes.remove(&handle_id) {
                                        let _ = child.kill();
                                    }
                                }
                                obj_ref.borrow_mut().fields.insert("hasexited".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            "start" => {
                                // Instance Start() — start process using StartInfo
                                let start_info = obj_ref.borrow().fields.get("startinfo").cloned().unwrap_or(Value::Nothing);
                                if let Value::Object(si_ref) = start_info {
                                    let filename = si_ref.borrow().fields.get("filename").map(|v| v.as_string()).unwrap_or_default();
                                    let args = si_ref.borrow().fields.get("arguments").map(|v| v.as_string()).unwrap_or_default();
                                    
                                    if !filename.is_empty() {
                                        let mut cmd = std::process::Command::new(&filename);
                                        if !args.is_empty() {
                                            for a in args.split_whitespace() {
                                                 cmd.arg(a);
                                            }
                                        }
                                        
                                        // Spawn process
                                        match cmd.spawn() {
                                            Ok(child) => {
                                                let handle_id = generate_runtime_id();
                                                obj_ref.borrow_mut().fields.insert("__handle".to_string(), Value::String(handle_id.clone()));
                                                obj_ref.borrow_mut().fields.insert("id".to_string(), Value::Integer(child.id() as i32));
                                                get_registry().lock().unwrap().processes.insert(handle_id, child);
                                                obj_ref.borrow_mut().fields.insert("hasexited".to_string(), Value::Boolean(false));
                                                return Ok(Value::Boolean(true));
                                            }
                                            Err(e) => return Err(RuntimeError::Custom(format!("Process.Start failed: {}", e))),
                                        }
                                    }
                                }
                                return Ok(Value::Boolean(false));
                            }
                            _ => {}
                        }
                    }

                    // StreamReader instance methods
                    if type_name == "StreamReader" {
                        // Check if this is a network-backed StreamReader
                        let socket_id = obj_ref.borrow().fields.get("__socket_id")
                            .and_then(|v| if let Value::Long(l) = v { Some(*l) } else { None });

                        if let Some(sid) = socket_id {
                            // Network-backed StreamReader
                            match method_name.as_str() {
                                "readline" => {
                                    if let Some(handle) = self.net_handles.get_mut(&sid) {
                                        match handle.tcp_read_line() {
                                            Ok(Some(line)) => return Ok(Value::String(line)),
                                            Ok(None) => return Ok(Value::Nothing), // EOF
                                            Err(e) => return Err(RuntimeError::Custom(format!("StreamReader.ReadLine: {}", e))),
                                        }
                                    }
                                    return Ok(Value::Nothing);
                                }
                                "readtoend" => {
                                    if let Some(handle) = self.net_handles.get_mut(&sid) {
                                        match handle.tcp_read_to_end() {
                                            Ok(data) => return Ok(Value::String(data)),
                                            Err(e) => return Err(RuntimeError::Custom(format!("StreamReader.ReadToEnd: {}", e))),
                                        }
                                    }
                                    return Ok(Value::String(String::new()));
                                }
                                "read" => {
                                    if let Some(handle) = self.net_handles.get_mut(&sid) {
                                        match handle.tcp_read_byte() {
                                            Ok(Some(b)) => return Ok(Value::Integer(b as i32)),
                                            Ok(None) => return Ok(Value::Integer(-1)),
                                            Err(e) => return Err(RuntimeError::Custom(format!("StreamReader.Read: {}", e))),
                                        }
                                    }
                                    return Ok(Value::Integer(-1));
                                }
                                "peek" => {
                                    // Peek is hard without buffering; just return -1 for network streams
                                    return Ok(Value::Integer(-1));
                                }
                                "close" | "dispose" => {
                                    obj_ref.borrow_mut().fields.insert("__closed".to_string(), Value::Boolean(true));
                                    return Ok(Value::Nothing);
                                }
                                _ => {}
                            }
                        } else {
                            // File-backed StreamReader (existing logic)
                            match method_name.as_str() {
                                "readline" => {
                                    let content = obj_ref.borrow().fields.get("__content").map(|v| v.as_string()).unwrap_or_default();
                                    let pos = obj_ref.borrow().fields.get("__position").map(|v| if let Value::Integer(i) = v { *i as usize } else { 0 }).unwrap_or(0);
                                    if pos >= content.len() {
                                        return Ok(Value::Nothing); // EOF
                                    }
                                    let remaining = &content[pos..];
                                    if let Some(nl) = remaining.find('\n') {
                                        let line = remaining[..nl].trim_end_matches('\r').to_string();
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Integer((pos + nl + 1) as i32));
                                        return Ok(Value::String(line));
                                    } else {
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Integer(content.len() as i32));
                                        return Ok(Value::String(remaining.to_string()));
                                    }
                                }
                                "readtoend" => {
                                    let content = obj_ref.borrow().fields.get("__content").map(|v| v.as_string()).unwrap_or_default();
                                    let pos = obj_ref.borrow().fields.get("__position").map(|v| if let Value::Integer(i) = v { *i as usize } else { 0 }).unwrap_or(0);
                                    obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Integer(content.len() as i32));
                                    if pos >= content.len() {
                                        return Ok(Value::String(String::new()));
                                    }
                                    return Ok(Value::String(content[pos..].to_string()));
                                }
                                "peek" => {
                                    let content = obj_ref.borrow().fields.get("__content").map(|v| v.as_string()).unwrap_or_default();
                                    let pos = obj_ref.borrow().fields.get("__position").map(|v| if let Value::Integer(i) = v { *i as usize } else { 0 }).unwrap_or(0);
                                    if pos >= content.len() { return Ok(Value::Integer(-1)); }
                                    return Ok(Value::Integer(content.as_bytes()[pos] as i32));
                                }
                                "close" | "dispose" => {
                                    obj_ref.borrow_mut().fields.insert("__closed".to_string(), Value::Boolean(true));
                                    return Ok(Value::Nothing);
                                }
                                _ => {}
                            }
                        }
                    }

                    // StreamWriter instance methods
                    if type_name == "StreamWriter" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                            .map(|arg| self.evaluate_expr(arg))
                            .collect();
                        let arg_values = arg_values?;

                        // Check if this is a network-backed StreamWriter
                        let socket_id = obj_ref.borrow().fields.get("__socket_id")
                            .and_then(|v| if let Value::Long(l) = v { Some(*l) } else { None });

                        if let Some(sid) = socket_id {
                            // Network-backed StreamWriter — write directly to socket
                            match method_name.as_str() {
                                "write" => {
                                    let text = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                    if let Some(handle) = self.net_handles.get_mut(&sid) {
                                        handle.tcp_write(text.as_bytes())
                                            .map_err(|e| RuntimeError::Custom(format!("StreamWriter.Write: {}", e)))?;
                                    }
                                    return Ok(Value::Nothing);
                                }
                                "writeline" => {
                                    let text = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                    if let Some(handle) = self.net_handles.get_mut(&sid) {
                                        handle.tcp_write_line(&text)
                                            .map_err(|e| RuntimeError::Custom(format!("StreamWriter.WriteLine: {}", e)))?;
                                    }
                                    return Ok(Value::Nothing);
                                }
                                "flush" => {
                                    if let Some(handle) = self.net_handles.get_mut(&sid) {
                                        handle.tcp_flush()
                                            .map_err(|e| RuntimeError::Custom(format!("StreamWriter.Flush: {}", e)))?;
                                    }
                                    return Ok(Value::Nothing);
                                }
                                "close" | "dispose" => {
                                    obj_ref.borrow_mut().fields.insert("__closed".to_string(), Value::Boolean(true));
                                    return Ok(Value::Nothing);
                                }
                                _ => {}
                            }
                        } else {
                            // File-backed StreamWriter (existing logic)
                            match method_name.as_str() {
                                "write" => {
                                    let text = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                    let mut buf = obj_ref.borrow().fields.get("__buffer").map(|v| v.as_string()).unwrap_or_default();
                                    buf.push_str(&text);
                                    obj_ref.borrow_mut().fields.insert("__buffer".to_string(), Value::String(buf));
                                    return Ok(Value::Nothing);
                                }
                                "writeline" => {
                                    let text = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                    let mut buf = obj_ref.borrow().fields.get("__buffer").map(|v| v.as_string()).unwrap_or_default();
                                    buf.push_str(&text);
                                    buf.push('\n');
                                    obj_ref.borrow_mut().fields.insert("__buffer".to_string(), Value::String(buf));
                                    return Ok(Value::Nothing);
                                }
                                "flush" => {
                                    let path = obj_ref.borrow().fields.get("__path").map(|v| v.as_string()).unwrap_or_default();
                                    let buf = obj_ref.borrow().fields.get("__buffer").map(|v| v.as_string()).unwrap_or_default();
                                    let append = obj_ref.borrow().fields.get("__append").map(|v| if let Value::Boolean(b) = v { *b } else { false }).unwrap_or(false);
                                    if append {
                                        use std::io::Write;
                                        let mut f = std::fs::OpenOptions::new().create(true).append(true).open(&path)
                                            .map_err(|e| RuntimeError::Custom(format!("StreamWriter.Flush: {}", e)))?;
                                        f.write_all(buf.as_bytes()).map_err(|e| RuntimeError::Custom(format!("StreamWriter.Flush: {}", e)))?;
                                    } else {
                                        std::fs::write(&path, &buf).map_err(|e| RuntimeError::Custom(format!("StreamWriter.Flush: {}", e)))?;
                                    }
                                    obj_ref.borrow_mut().fields.insert("__buffer".to_string(), Value::String(String::new()));
                                    obj_ref.borrow_mut().fields.insert("__append".to_string(), Value::Boolean(true));
                                    return Ok(Value::Nothing);
                                }
                                "close" | "dispose" => {
                                    let path = obj_ref.borrow().fields.get("__path").map(|v| v.as_string()).unwrap_or_default();
                                    let buf = obj_ref.borrow().fields.get("__buffer").map(|v| v.as_string()).unwrap_or_default();
                                    let append = obj_ref.borrow().fields.get("__append").map(|v| if let Value::Boolean(b) = v { *b } else { false }).unwrap_or(false);
                                    if !buf.is_empty() {
                                        if append {
                                            use std::io::Write;
                                            let mut f = std::fs::OpenOptions::new().create(true).append(true).open(&path)
                                                .map_err(|e| RuntimeError::Custom(format!("StreamWriter.Close: {}", e)))?;
                                            f.write_all(buf.as_bytes()).map_err(|e| RuntimeError::Custom(format!("StreamWriter.Close: {}", e)))?;
                                        } else {
                                            std::fs::write(&path, &buf).map_err(|e| RuntimeError::Custom(format!("StreamWriter.Close: {}", e)))?;
                                        }
                                    }
                                    obj_ref.borrow_mut().fields.insert("__buffer".to_string(), Value::String(String::new()));
                                    obj_ref.borrow_mut().fields.insert("__closed".to_string(), Value::Boolean(true));
                                    return Ok(Value::Nothing);
                                }
                                _ => {}
                            }
                        }
                    }

                    // Regex instance methods
                    if type_name == "Regex" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                            .map(|arg| self.evaluate_expr(arg))
                            .collect();
                        let arg_values = arg_values?;
                        let pattern = obj_ref.borrow().fields.get("__pattern").map(|v| v.as_string()).unwrap_or_default();
                        let options = obj_ref.borrow().fields.get("__options").map(|v| if let Value::Integer(i) = v { *i } else { 0 }).unwrap_or(0);
                        let full_pattern = if options & 1 != 0 {
                            format!("(?i){}", pattern)
                        } else {
                            pattern.clone()
                        };
                        match method_name.as_str() {
                            "ismatch" => {
                                let input = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                match regex::Regex::new(&full_pattern) {
                                    Ok(re) => return Ok(Value::Boolean(re.is_match(&input))),
                                    Err(e) => return Err(RuntimeError::Custom(format!("Invalid regex: {}", e))),
                                }
                            }
                            "match" => {
                                let input = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                let args_for_fn = vec![Value::String(input), Value::String(full_pattern)];
                                return crate::builtins::text_fns::regex_match_fn(&args_for_fn);
                            }
                            "matches" => {
                                let input = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                let args_for_fn = vec![Value::String(input), Value::String(full_pattern)];
                                return crate::builtins::text_fns::regex_matches_fn(&args_for_fn);
                            }
                            "replace" => {
                                let input = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                if arg_values.len() >= 3 {
                                    // Static-style: regex.Replace(input, pattern, replacement[, options])
                                    let pat = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                                    let replacement = arg_values.get(2).map(|v| v.as_string()).unwrap_or_default();
                                    let opts = arg_values.get(3).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                    let effective = if opts & 1 != 0 { format!("(?i){}", pat) } else { pat };
                                    let args_for_fn = vec![Value::String(input), Value::String(effective), Value::String(replacement)];
                                    return crate::builtins::text_fns::regex_replace_fn(&args_for_fn);
                                } else {
                                    // Instance: regex.Replace(input, replacement)
                                    let replacement = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                                    let args_for_fn = vec![Value::String(input), Value::String(full_pattern), Value::String(replacement)];
                                    return crate::builtins::text_fns::regex_replace_fn(&args_for_fn);
                                }
                            }
                            "split" => {
                                let input = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                let args_for_fn = vec![Value::String(input), Value::String(full_pattern)];
                                return crate::builtins::text_fns::regex_split_fn(&args_for_fn);
                            }
                            _ => {}
                        }
                    }

                    // ===== Task instance methods =====
                    if type_name == "Task" {
                        match method_name.as_str() {
                            "wait" => { return Ok(Value::Nothing); } // Already completed
                            "getawaiter" => { return Ok(obj_val.clone()); } // Return self
                            "getresult" => {
                                return Ok(obj_ref.borrow().fields.get("result").cloned().unwrap_or(Value::Nothing));
                            }
                            "continueWith" | "continuewith" => {
                                let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                                let arg_values = arg_values?;
                                if let Some(lambda) = arg_values.get(0) {
                                    let result = self.call_lambda(lambda.clone(), &[obj_val.clone()])?;
                                    let mut fields = std::collections::HashMap::new();
                                    fields.insert("__type".to_string(), Value::String("Task".to_string()));
                                    fields.insert("result".to_string(), result);
                                    fields.insert("iscompleted".to_string(), Value::Boolean(true));
                                    fields.insert("isfaulted".to_string(), Value::Boolean(false));
                                    fields.insert("iscanceled".to_string(), Value::Boolean(false));
                                    fields.insert("status".to_string(), Value::String("RanToCompletion".to_string()));
                                    let obj = crate::value::ObjectData { class_name: "Task".to_string(), fields };
                                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                }
                                return Ok(Value::Nothing);
                            }
                            "configureawait" => { return Ok(obj_val.clone()); } // No-op, return self
                            _ => {}
                        }
                    }

                    // ===== Mutex instance methods =====
                    if type_name == "Mutex" {
                        match method_name.as_str() {
                            "waitone" => {
                                obj_ref.borrow_mut().fields.insert("__owned".to_string(), Value::Boolean(true));
                                return Ok(Value::Boolean(true));
                            }
                            "releasemutex" => {
                                obj_ref.borrow_mut().fields.insert("__owned".to_string(), Value::Boolean(false));
                                return Ok(Value::Nothing);
                            }
                            "close" | "dispose" => { return Ok(Value::Nothing); }
                            _ => {}
                        }
                    }

                    // ===== Semaphore instance methods =====
                    if type_name == "Semaphore" {
                        match method_name.as_str() {
                            "wait" | "waitone" => {
                                let count = obj_ref.borrow().fields.get("__count").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                if count > 0 {
                                    obj_ref.borrow_mut().fields.insert("__count".to_string(), Value::Integer(count - 1));
                                    obj_ref.borrow_mut().fields.insert("currentcount".to_string(), Value::Integer(count - 1));
                                }
                                return Ok(Value::Boolean(true));
                            }
                            "release" => {
                                let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                                let arg_values = arg_values?;
                                let release_count = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(1);
                                let count = obj_ref.borrow().fields.get("__count").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                let max = obj_ref.borrow().fields.get("__max").and_then(|v| v.as_integer().ok()).unwrap_or(i32::MAX);
                                let new_count = (count + release_count).min(max);
                                let prev = count;
                                obj_ref.borrow_mut().fields.insert("__count".to_string(), Value::Integer(new_count));
                                obj_ref.borrow_mut().fields.insert("currentcount".to_string(), Value::Integer(new_count));
                                return Ok(Value::Integer(prev));
                            }
                            "close" | "dispose" => { return Ok(Value::Nothing); }
                            _ => {}
                        }
                    }

                    // ===== Timer instance methods =====
                    if type_name == "Timer" {
                        match method_name.as_str() {
                            "start" => {
                                obj_ref.borrow_mut().fields.insert("enabled".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            "stop" => {
                                obj_ref.borrow_mut().fields.insert("enabled".to_string(), Value::Boolean(false));
                                return Ok(Value::Nothing);
                            }
                            "close" | "dispose" => {
                                obj_ref.borrow_mut().fields.insert("enabled".to_string(), Value::Boolean(false));
                                return Ok(Value::Nothing);
                            }
                            _ => {}
                        }
                    }

                    // ===== FileStream instance methods =====
                    if type_name == "FileStream" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                        let arg_values = arg_values?;
                        match method_name.as_str() {
                            "read" => {
                                // Read(buffer, offset, count) -> int
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    let count = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(arr.len() as i32) as usize;
                                    let end = (pos + count).min(arr.len());
                                    let bytes_read = end - pos;
                                    obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long(end as i64));
                                    obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Long(end as i64));
                                    return Ok(Value::Integer(bytes_read as i32));
                                }
                                return Ok(Value::Integer(0));
                            }
                            "readbyte" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    if pos < arr.len() {
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long((pos + 1) as i64));
                                        obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Long((pos + 1) as i64));
                                        return Ok(arr[pos].clone());
                                    }
                                }
                                return Ok(Value::Integer(-1));
                            }
                            "write" => {
                                // Write(buffer, offset, count)
                                if let Some(Value::Array(buf)) = arg_values.get(0) {
                                    let offset = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                                    let count = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(buf.len() as i32) as usize;
                                    let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                    if let Value::Array(mut arr) = data {
                                        for i in offset..(offset + count).min(buf.len()) {
                                            arr.push(buf[i].clone());
                                        }
                                        let len = arr.len() as i64;
                                        obj_ref.borrow_mut().fields.insert("__data".to_string(), Value::Array(arr));
                                        obj_ref.borrow_mut().fields.insert("length".to_string(), Value::Long(len));
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long(len));
                                        obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Long(len));
                                    }
                                }
                                return Ok(Value::Nothing);
                            }
                            "writebyte" => {
                                let byte_val = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                if let Value::Array(mut arr) = data {
                                    arr.push(Value::Integer(byte_val));
                                    let len = arr.len() as i64;
                                    obj_ref.borrow_mut().fields.insert("__data".to_string(), Value::Array(arr));
                                    obj_ref.borrow_mut().fields.insert("length".to_string(), Value::Long(len));
                                }
                                return Ok(Value::Nothing);
                            }
                            "seek" => {
                                let offset = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as i64;
                                let origin = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0); // 0=Begin, 1=Current, 2=End
                                let len = obj_ref.borrow().fields.get("length").and_then(|v| if let Value::Long(l) = v { Some(*l) } else { None }).unwrap_or(0);
                                let cur_pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l) } else { None }).unwrap_or(0);
                                let new_pos = match origin {
                                    0 => offset,
                                    1 => cur_pos + offset,
                                    2 => len + offset,
                                    _ => offset,
                                }.max(0);
                                obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long(new_pos));
                                obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Long(new_pos));
                                return Ok(Value::Long(new_pos));
                            }
                            "flush" => { return Ok(Value::Nothing); }
                            "close" | "dispose" => {
                                // Write data to file before closing
                                let path = obj_ref.borrow().fields.get("__path").map(|v| v.as_string()).unwrap_or_default();
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                if let Value::Array(ref arr) = data {
                                    let bytes: Vec<u8> = arr.iter().map(|v| v.as_integer().unwrap_or(0) as u8).collect();
                                    let _ = std::fs::write(&path, &bytes);
                                }
                                obj_ref.borrow_mut().fields.insert("__closed".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            "toarray" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                return Ok(data);
                            }
                            _ => {}
                        }
                    }

                    // ===== MemoryStream instance methods =====
                    if type_name == "MemoryStream" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                        let arg_values = arg_values?;
                        match method_name.as_str() {
                            "read" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    let count = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(arr.len() as i32) as usize;
                                    let end = (pos + count).min(arr.len());
                                    let bytes_read = end - pos;
                                    obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Long(end as i64));
                                    return Ok(Value::Integer(bytes_read as i32));
                                }
                                return Ok(Value::Integer(0));
                            }
                            "readbyte" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    if pos < arr.len() {
                                        obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Long((pos + 1) as i64));
                                        return Ok(arr[pos].clone());
                                    }
                                }
                                return Ok(Value::Integer(-1));
                            }
                            "write" => {
                                if let Some(Value::Array(buf)) = arg_values.get(0) {
                                    let offset = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                                    let count = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(buf.len() as i32) as usize;
                                    let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                    if let Value::Array(mut arr) = data {
                                        for i in offset..(offset + count).min(buf.len()) {
                                            arr.push(buf[i].clone());
                                        }
                                        let len = arr.len() as i64;
                                        obj_ref.borrow_mut().fields.insert("__data".to_string(), Value::Array(arr));
                                        obj_ref.borrow_mut().fields.insert("length".to_string(), Value::Long(len));
                                    }
                                }
                                return Ok(Value::Nothing);
                            }
                            "writebyte" => {
                                let byte_val = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                if let Value::Array(mut arr) = data {
                                    arr.push(Value::Integer(byte_val));
                                    let len = arr.len() as i64;
                                    obj_ref.borrow_mut().fields.insert("__data".to_string(), Value::Array(arr));
                                    obj_ref.borrow_mut().fields.insert("length".to_string(), Value::Long(len));
                                }
                                return Ok(Value::Nothing);
                            }
                            "seek" => {
                                let offset = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as i64;
                                let origin = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                let len = obj_ref.borrow().fields.get("length").and_then(|v| if let Value::Long(l) = v { Some(*l) } else { None }).unwrap_or(0);
                                let cur = obj_ref.borrow().fields.get("position").and_then(|v| if let Value::Long(l) = v { Some(*l) } else { None }).unwrap_or(0);
                                let new_pos = match origin {
                                    0 => offset, 1 => cur + offset, 2 => len + offset, _ => offset,
                                }.max(0);
                                obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Long(new_pos));
                                return Ok(Value::Long(new_pos));
                            }
                            "toarray" | "getbuffer" => {
                                return Ok(obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new())));
                            }
                            "setlength" => {
                                let new_len = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                if let Value::Array(mut arr) = data {
                                    arr.resize(new_len, Value::Integer(0));
                                    obj_ref.borrow_mut().fields.insert("__data".to_string(), Value::Array(arr));
                                    obj_ref.borrow_mut().fields.insert("length".to_string(), Value::Long(new_len as i64));
                                }
                                return Ok(Value::Nothing);
                            }
                            "flush" | "close" | "dispose" => { return Ok(Value::Nothing); }
                            _ => {}
                        }
                    }

                    // ===== BinaryReader instance methods =====
                    if type_name == "BinaryReader" {
                        match method_name.as_str() {
                            "readbyte" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    if pos < arr.len() {
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long((pos + 1) as i64));
                                        return Ok(arr[pos].clone());
                                    }
                                }
                                return Err(RuntimeError::Custom("BinaryReader: end of stream".to_string()));
                            }
                            "readint16" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    if pos + 2 <= arr.len() {
                                        let b0 = arr[pos].as_integer().unwrap_or(0) as u8;
                                        let b1 = arr[pos+1].as_integer().unwrap_or(0) as u8;
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long((pos + 2) as i64));
                                        return Ok(Value::Integer(i16::from_le_bytes([b0, b1]) as i32));
                                    }
                                }
                                return Err(RuntimeError::Custom("BinaryReader: end of stream".to_string()));
                            }
                            "readint32" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    if pos + 4 <= arr.len() {
                                        let bytes: Vec<u8> = arr[pos..pos+4].iter().map(|v| v.as_integer().unwrap_or(0) as u8).collect();
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long((pos + 4) as i64));
                                        return Ok(Value::Integer(i32::from_le_bytes([bytes[0],bytes[1],bytes[2],bytes[3]])));
                                    }
                                }
                                return Err(RuntimeError::Custom("BinaryReader: end of stream".to_string()));
                            }
                            "readint64" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    if pos + 8 <= arr.len() {
                                        let bytes: Vec<u8> = arr[pos..pos+8].iter().map(|v| v.as_integer().unwrap_or(0) as u8).collect();
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long((pos + 8) as i64));
                                        return Ok(Value::Long(i64::from_le_bytes([bytes[0],bytes[1],bytes[2],bytes[3],bytes[4],bytes[5],bytes[6],bytes[7]])));
                                    }
                                }
                                return Err(RuntimeError::Custom("BinaryReader: end of stream".to_string()));
                            }
                            "readdouble" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    if pos + 8 <= arr.len() {
                                        let bytes: Vec<u8> = arr[pos..pos+8].iter().map(|v| v.as_integer().unwrap_or(0) as u8).collect();
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long((pos + 8) as i64));
                                        return Ok(Value::Double(f64::from_le_bytes([bytes[0],bytes[1],bytes[2],bytes[3],bytes[4],bytes[5],bytes[6],bytes[7]])));
                                    }
                                }
                                return Err(RuntimeError::Custom("BinaryReader: end of stream".to_string()));
                            }
                            "readboolean" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    if pos < arr.len() {
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long((pos + 1) as i64));
                                        return Ok(Value::Boolean(arr[pos].as_integer().unwrap_or(0) != 0));
                                    }
                                }
                                return Err(RuntimeError::Custom("BinaryReader: end of stream".to_string()));
                            }
                            "readstring" => {
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    // Read 7-bit encoded length, then that many bytes
                                    if pos < arr.len() {
                                        let str_len = arr[pos].as_integer().unwrap_or(0) as usize;
                                        let start = pos + 1;
                                        let end = (start + str_len).min(arr.len());
                                        let bytes: Vec<u8> = arr[start..end].iter().map(|v| v.as_integer().unwrap_or(0) as u8).collect();
                                        obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long(end as i64));
                                        return Ok(Value::String(String::from_utf8_lossy(&bytes).to_string()));
                                    }
                                }
                                return Err(RuntimeError::Custom("BinaryReader: end of stream".to_string()));
                            }
                            "readbytes" => {
                                let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                                let arg_values = arg_values?;
                                let count = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                let pos = obj_ref.borrow().fields.get("__position").and_then(|v| if let Value::Long(l) = v { Some(*l as usize) } else { None }).unwrap_or(0);
                                if let Value::Array(ref arr) = data {
                                    let end = (pos + count).min(arr.len());
                                    let result: Vec<Value> = arr[pos..end].to_vec();
                                    obj_ref.borrow_mut().fields.insert("__position".to_string(), Value::Long(end as i64));
                                    return Ok(Value::Array(result));
                                }
                                return Ok(Value::Array(Vec::new()));
                            }
                            "close" | "dispose" => {
                                obj_ref.borrow_mut().fields.insert("__closed".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            _ => {}
                        }
                    }

                    // ===== BinaryWriter instance methods =====
                    if type_name == "BinaryWriter" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                        let arg_values = arg_values?;
                        match method_name.as_str() {
                            "write" => {
                                let val = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                                let bytes: Vec<Value> = match &val {
                                    Value::Integer(i) => i.to_le_bytes().iter().map(|b| Value::Integer(*b as i32)).collect(),
                                    Value::Long(l) => l.to_le_bytes().iter().map(|b| Value::Integer(*b as i32)).collect(),
                                    Value::Double(d) => d.to_le_bytes().iter().map(|b| Value::Integer(*b as i32)).collect(),
                                    Value::Single(f) => f.to_le_bytes().iter().map(|b| Value::Integer(*b as i32)).collect(),
                                    Value::Boolean(b) => vec![Value::Integer(if *b { 1 } else { 0 })],
                                    Value::String(s) => {
                                        let mut v = vec![Value::Integer(s.len() as i32)]; // length prefix
                                        for b in s.bytes() { v.push(Value::Integer(b as i32)); }
                                        v
                                    }
                                    Value::Array(arr) => arr.clone(),
                                    _ => vec![],
                                };
                                let data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                if let Value::Array(mut arr) = data {
                                    arr.extend(bytes);
                                    obj_ref.borrow_mut().fields.insert("__data".to_string(), Value::Array(arr));
                                }
                                // Also write to underlying stream if present
                                if let Some(Value::Object(stream_ref)) = obj_ref.borrow().fields.get("__stream") {
                                    let sdata = stream_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                    if let Value::Array(mut sarr) = sdata {
                                        let writer_data = obj_ref.borrow().fields.get("__data").cloned().unwrap_or(Value::Array(Vec::new()));
                                        if let Value::Array(wd) = writer_data { sarr = wd; }
                                        let len = sarr.len() as i64;
                                        stream_ref.borrow_mut().fields.insert("__data".to_string(), Value::Array(sarr));
                                        stream_ref.borrow_mut().fields.insert("length".to_string(), Value::Long(len));
                                    }
                                }
                                return Ok(Value::Nothing);
                            }
                            "flush" => { return Ok(Value::Nothing); }
                            "close" | "dispose" => {
                                obj_ref.borrow_mut().fields.insert("__closed".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            _ => {}
                        }
                    }

                    // ===== TcpClient instance methods =====
                    if type_name == "TcpClient" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                        let arg_values = arg_values?;
                        match method_name.as_str() {
                            "connect" => {
                                let host = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                let port = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                // Remove old handle if any
                                if let Some(Value::Long(old_id)) = obj_ref.borrow().fields.get("__socket_id") {
                                    if *old_id > 0 { self.remove_net_handle(*old_id); }
                                }
                                match crate::builtins::networking::NetHandle::connect_tcp(&host, port, None) {
                                    Ok(handle) => {
                                        let id = self.alloc_net_handle(handle);
                                        let mut b = obj_ref.borrow_mut();
                                        b.fields.insert("__host".to_string(), Value::String(host));
                                        b.fields.insert("__port".to_string(), Value::Integer(port));
                                        b.fields.insert("__socket_id".to_string(), Value::Long(id));
                                        b.fields.insert("connected".to_string(), Value::Boolean(true));
                                    }
                                    Err(e) => {
                                        return Err(RuntimeError::Custom(format!("TcpClient.Connect failed: {}", e)));
                                    }
                                }
                                return Ok(Value::Nothing);
                            }
                            "getstream" => {
                                let socket_id = obj_ref.borrow().fields.get("__socket_id")
                                    .and_then(|v| if let Value::Long(id) = v { Some(*id) } else { None })
                                    .unwrap_or(0);
                                let mut fields = std::collections::HashMap::new();
                                fields.insert("__type".to_string(), Value::String("NetworkStream".to_string()));
                                fields.insert("__socket_id".to_string(), Value::Long(socket_id));
                                fields.insert("canread".to_string(), Value::Boolean(true));
                                fields.insert("canwrite".to_string(), Value::Boolean(true));
                                let obj = crate::value::ObjectData { class_name: "NetworkStream".to_string(), fields };
                                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                            }
                            "close" | "dispose" => {
                                if let Some(Value::Long(id)) = obj_ref.borrow().fields.get("__socket_id") {
                                    if *id > 0 {
                                        let id_val = *id;
                                        if let Some(handle) = self.net_handles.get_mut(&id_val) {
                                            let _ = handle.tcp_shutdown();
                                        }
                                        self.remove_net_handle(id_val);
                                    }
                                }
                                let mut b = obj_ref.borrow_mut();
                                b.fields.insert("connected".to_string(), Value::Boolean(false));
                                b.fields.insert("__socket_id".to_string(), Value::Long(0));
                                return Ok(Value::Nothing);
                            }
                            _ => {}
                        }
                    }

                    // ===== NetworkStream instance methods =====
                    if type_name == "NetworkStream" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                        let arg_values = arg_values?;
                        let socket_id = obj_ref.borrow().fields.get("__socket_id")
                            .and_then(|v| if let Value::Long(id) = v { Some(*id) } else { None })
                            .unwrap_or(0);
                        match method_name.as_str() {
                            "read" => {
                                // Read(buffer, offset, count) → returns bytes read
                                let count = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(4096) as usize;
                                if let Some(handle) = self.net_handles.get_mut(&socket_id) {
                                    let mut buf = vec![0u8; count];
                                    match handle.tcp_read(&mut buf) {
                                        Ok(n) => {
                                            buf.truncate(n);
                                            // Update the buffer array arg if provided
                                            return Ok(Value::Integer(n as i32));
                                        }
                                        Err(e) => return Err(RuntimeError::Custom(format!("NetworkStream.Read: {}", e))),
                                    }
                                }
                                return Ok(Value::Integer(0));
                            }
                            "readbyte" => {
                                if let Some(handle) = self.net_handles.get_mut(&socket_id) {
                                    match handle.tcp_read_byte() {
                                        Ok(Some(b)) => return Ok(Value::Integer(b as i32)),
                                        Ok(None) => return Ok(Value::Integer(-1)), // EOF
                                        Err(e) => return Err(RuntimeError::Custom(format!("NetworkStream.ReadByte: {}", e))),
                                    }
                                }
                                return Ok(Value::Integer(-1));
                            }
                            "write" => {
                                // Write(buffer, offset, count)
                                let data: Vec<u8> = if let Some(Value::Array(arr)) = arg_values.get(0) {
                                    arr.iter().map(|v| v.as_integer().unwrap_or(0) as u8).collect()
                                } else if let Some(val) = arg_values.get(0) {
                                    val.as_string().into_bytes()
                                } else {
                                    Vec::new()
                                };
                                if let Some(handle) = self.net_handles.get_mut(&socket_id) {
                                    match handle.tcp_write(&data) {
                                        Ok(_) => {}
                                        Err(e) => return Err(RuntimeError::Custom(format!("NetworkStream.Write: {}", e))),
                                    }
                                }
                                return Ok(Value::Nothing);
                            }
                            "writebyte" => {
                                let b = arg_values.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as u8;
                                if let Some(handle) = self.net_handles.get_mut(&socket_id) {
                                    let _ = handle.tcp_write(&[b]);
                                }
                                return Ok(Value::Nothing);
                            }
                            "flush" => {
                                if let Some(handle) = self.net_handles.get_mut(&socket_id) {
                                    let _ = handle.tcp_flush();
                                }
                                return Ok(Value::Nothing);
                            }
                            "close" | "dispose" => {
                                if socket_id > 0 {
                                    if let Some(handle) = self.net_handles.get_mut(&socket_id) {
                                        let _ = handle.tcp_shutdown();
                                    }
                                }
                                return Ok(Value::Nothing);
                            }
                            _ => {}
                        }
                    }

                    // ===== TcpListener instance methods =====
                    if type_name == "TcpListener" {
                        match method_name.as_str() {
                            "start" => {
                                let addr = obj_ref.borrow().fields.get("__address")
                                    .map(|v| v.as_string()).unwrap_or_else(|| "0.0.0.0".to_string());
                                let port = obj_ref.borrow().fields.get("__port")
                                    .and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                // Remove old handle if any
                                if let Some(Value::Long(old_id)) = obj_ref.borrow().fields.get("__socket_id") {
                                    if *old_id > 0 { self.remove_net_handle(*old_id); }
                                }
                                match crate::builtins::networking::NetHandle::bind_listener(&addr, port) {
                                    Ok(handle) => {
                                        let id = self.alloc_net_handle(handle);
                                        let mut b = obj_ref.borrow_mut();
                                        b.fields.insert("__socket_id".to_string(), Value::Long(id));
                                        b.fields.insert("__active".to_string(), Value::Boolean(true));
                                    }
                                    Err(e) => {
                                        return Err(RuntimeError::Custom(format!("TcpListener.Start failed: {}", e)));
                                    }
                                }
                                return Ok(Value::Nothing);
                            }
                            "stop" => {
                                if let Some(Value::Long(id)) = obj_ref.borrow().fields.get("__socket_id") {
                                    if *id > 0 { self.remove_net_handle(*id); }
                                }
                                let mut b = obj_ref.borrow_mut();
                                b.fields.insert("__active".to_string(), Value::Boolean(false));
                                b.fields.insert("__socket_id".to_string(), Value::Long(0));
                                return Ok(Value::Nothing);
                            }
                            "accepttcpclient" => {
                                let socket_id = obj_ref.borrow().fields.get("__socket_id")
                                    .and_then(|v| if let Value::Long(id) = v { Some(*id) } else { None })
                                    .unwrap_or(0);
                                if let Some(handle) = self.net_handles.get(&socket_id) {
                                    match handle.listener_accept() {
                                        Ok((stream, addr)) => {
                                            let peer_host = addr.split(':').next().unwrap_or("").to_string();
                                            let peer_port: i32 = addr.split(':').last()
                                                .and_then(|s| s.parse().ok()).unwrap_or(0);
                                            match crate::builtins::networking::NetHandle::from_tcp_stream(stream) {
                                                Ok(client_handle) => {
                                                    let client_id = self.alloc_net_handle(client_handle);
                                                    let mut fields = std::collections::HashMap::new();
                                                    fields.insert("__type".to_string(), Value::String("TcpClient".to_string()));
                                                    fields.insert("connected".to_string(), Value::Boolean(true));
                                                    fields.insert("__host".to_string(), Value::String(peer_host));
                                                    fields.insert("__port".to_string(), Value::Integer(peer_port));
                                                    fields.insert("__socket_id".to_string(), Value::Long(client_id));
                                                    fields.insert("receivebuffersize".to_string(), Value::Integer(8192));
                                                    fields.insert("sendbuffersize".to_string(), Value::Integer(8192));
                                                    let obj = crate::value::ObjectData { class_name: "TcpClient".to_string(), fields };
                                                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                                }
                                                Err(e) => return Err(RuntimeError::Custom(format!("AcceptTcpClient: {}", e))),
                                            }
                                        }
                                        Err(e) => return Err(RuntimeError::Custom(format!("AcceptTcpClient: {}", e))),
                                    }
                                }
                                return Err(RuntimeError::Custom("TcpListener not started".to_string()));
                            }
                            "pending" => {
                                let socket_id = obj_ref.borrow().fields.get("__socket_id")
                                    .and_then(|v| if let Value::Long(id) = v { Some(*id) } else { None })
                                    .unwrap_or(0);
                                if let Some(handle) = self.net_handles.get(&socket_id) {
                                    match handle.listener_pending() {
                                        Ok(b) => return Ok(Value::Boolean(b)),
                                        Err(_) => return Ok(Value::Boolean(false)),
                                    }
                                }
                                return Ok(Value::Boolean(false));
                            }
                            _ => {}
                        }
                    }

                    // ===== UdpClient instance methods =====
                    if type_name == "UdpClient" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                        let arg_values = arg_values?;
                        let socket_id = obj_ref.borrow().fields.get("__socket_id")
                            .and_then(|v| if let Value::Long(id) = v { Some(*id) } else { None })
                            .unwrap_or(0);
                        match method_name.as_str() {
                            "send" => {
                                // Send(bytes, length, hostname, port)
                                let data: Vec<u8> = if let Some(Value::Array(arr)) = arg_values.get(0) {
                                    arr.iter().map(|v| v.as_integer().unwrap_or(0) as u8).collect()
                                } else {
                                    Vec::new()
                                };
                                let hostname = arg_values.get(2).map(|v| v.as_string()).unwrap_or_default();
                                let port = arg_values.get(3).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                if let Some(handle) = self.net_handles.get(&socket_id) {
                                    if !hostname.is_empty() {
                                        match handle.udp_send_to(&data, &hostname, port) {
                                            Ok(n) => return Ok(Value::Integer(n as i32)),
                                            Err(e) => return Err(RuntimeError::Custom(format!("UdpClient.Send: {}", e))),
                                        }
                                    } else {
                                        match handle.udp_send(&data) {
                                            Ok(n) => return Ok(Value::Integer(n as i32)),
                                            Err(e) => return Err(RuntimeError::Custom(format!("UdpClient.Send: {}", e))),
                                        }
                                    }
                                }
                                return Ok(Value::Integer(0));
                            }
                            "receive" => {
                                // Receive(ByRef remoteEP) → returns byte array
                                let buf_size = obj_ref.borrow().fields.get("receivebuffersize")
                                    .and_then(|v| v.as_integer().ok()).unwrap_or(8192) as usize;
                                if let Some(handle) = self.net_handles.get(&socket_id) {
                                    match handle.udp_recv(buf_size) {
                                        Ok((data, _addr)) => {
                                            let arr: Vec<Value> = data.into_iter().map(|b| Value::Integer(b as i32)).collect();
                                            return Ok(Value::Array(arr));
                                        }
                                        Err(e) => return Err(RuntimeError::Custom(format!("UdpClient.Receive: {}", e))),
                                    }
                                }
                                return Ok(Value::Array(Vec::new()));
                            }
                            "connect" => {
                                let host = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                                let port = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                if let Some(handle) = self.net_handles.get(&socket_id) {
                                    if let Err(e) = handle.udp_connect(&host, port) {
                                        return Err(RuntimeError::Custom(format!("UdpClient.Connect: {}", e)));
                                    }
                                }
                                return Ok(Value::Nothing);
                            }
                            "close" | "dispose" => {
                                if socket_id > 0 {
                                    self.remove_net_handle(socket_id);
                                }
                                return Ok(Value::Nothing);
                            }
                            _ => {}
                        }
                    }

                    // ===== SmtpClient instance methods =====
                    if type_name == "SmtpClient" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                        let arg_values = arg_values?;
                        match method_name.as_str() {
                            "send" => {
                                // SmtpClient.Send(MailMessage) — use curl SMTP
                                if let Some(Value::Object(msg_ref)) = arg_values.get(0) {
                                    let msg = msg_ref.borrow();
                                    let from = msg.fields.get("from").map(|v| v.as_string()).unwrap_or_default();
                                    let to = msg.fields.get("to").map(|v| v.as_string()).unwrap_or_default();
                                    let subject = msg.fields.get("subject").map(|v| v.as_string()).unwrap_or_default();
                                    let body = msg.fields.get("body").map(|v| v.as_string()).unwrap_or_default();
                                    let is_html = msg.fields.get("isbodyhtml").and_then(|v| v.as_bool().ok()).unwrap_or(false);
                                    // CC and BCC from MailMessage
                                    let cc = msg.fields.get("cc").map(|v| v.as_string()).unwrap_or_default();
                                    let bcc = msg.fields.get("bcc").map(|v| v.as_string()).unwrap_or_default();

                                    let host = obj_ref.borrow().fields.get("host").map(|v| v.as_string()).unwrap_or_default();
                                    let port = obj_ref.borrow().fields.get("port").and_then(|v| v.as_integer().ok()).unwrap_or(25);
                                    let enable_ssl = obj_ref.borrow().fields.get("enablessl").and_then(|v| v.as_bool().ok()).unwrap_or(false);

                                    // Credentials (NetworkCredential object or username/password strings)
                                    let (cred_user, cred_pass) = {
                                        let obj = obj_ref.borrow();
                                        if let Some(Value::Object(cred_ref)) = obj.fields.get("credentials") {
                                            let cred = cred_ref.borrow();
                                            let u = cred.fields.get("username").map(|v| v.as_string()).unwrap_or_default();
                                            let p = cred.fields.get("password").map(|v| v.as_string()).unwrap_or_default();
                                            (u, p)
                                        } else {
                                            (String::new(), String::new())
                                        }
                                    };

                                    // Build SMTP URL — use smtps:// for SSL on port 465, else smtp:// with --ssl for STARTTLS
                                    let smtp_url = if enable_ssl && port == 465 {
                                        format!("smtps://{}:{}", host, port)
                                    } else {
                                        format!("smtp://{}:{}", host, port)
                                    };

                                    let mut curl_args: Vec<String> = vec![
                                        "--mail-from".to_string(), from.clone(),
                                        "--mail-rcpt".to_string(), to.clone(),
                                    ];

                                    // Add CC recipients
                                    if !cc.is_empty() {
                                        for addr in cc.split(';').map(|s| s.trim()).filter(|s| !s.is_empty()) {
                                            curl_args.push("--mail-rcpt".to_string());
                                            curl_args.push(addr.to_string());
                                        }
                                    }
                                    // Add BCC recipients
                                    if !bcc.is_empty() {
                                        for addr in bcc.split(';').map(|s| s.trim()).filter(|s| !s.is_empty()) {
                                            curl_args.push("--mail-rcpt".to_string());
                                            curl_args.push(addr.to_string());
                                        }
                                    }

                                    curl_args.push("--url".to_string());
                                    curl_args.push(smtp_url);

                                    // SSL/TLS: STARTTLS for non-465 ports
                                    if enable_ssl && port != 465 {
                                        curl_args.push("--ssl-reqd".to_string());
                                    }

                                    // Credentials
                                    if !cred_user.is_empty() {
                                        curl_args.push("--user".to_string());
                                        curl_args.push(format!("{}:{}", cred_user, cred_pass));
                                    }

                                    curl_args.push("-T".to_string());
                                    curl_args.push("-".to_string());

                                    // Build email body with headers
                                    let content_type = if is_html { "text/html" } else { "text/plain" };
                                    let mut email_body = format!(
                                        "From: {}\r\nTo: {}\r\nSubject: {}\r\nContent-Type: {}; charset=utf-8\r\n",
                                        from, to, subject, content_type
                                    );
                                    if !cc.is_empty() {
                                        email_body.push_str(&format!("Cc: {}\r\n", cc));
                                    }
                                    email_body.push_str(&format!("\r\n{}", body));

                                    let curl_args_str: Vec<&str> = curl_args.iter().map(|s| s.as_str()).collect();
                                    let status = std::process::Command::new("curl")
                                        .args(&curl_args_str)
                                        .stdin(std::process::Stdio::piped())
                                        .stdout(std::process::Stdio::null())
                                        .stderr(std::process::Stdio::piped())
                                        .spawn()
                                        .and_then(|mut child| {
                                            if let Some(ref mut stdin) = child.stdin {
                                                use std::io::Write;
                                                let _ = stdin.write_all(email_body.as_bytes());
                                            }
                                            child.wait()
                                        });
                                    match status {
                                        Ok(s) if s.success() => return Ok(Value::Nothing),
                                        Ok(s) => return Err(RuntimeError::Custom(format!("SmtpClient.Send failed with exit code {}", s.code().unwrap_or(-1)))),
                                        Err(e) => return Err(RuntimeError::Custom(format!("SmtpClient.Send failed: {}", e))),
                                    }
                                }
                                return Ok(Value::Nothing);
                            }
                            "dispose" | "close" => { return Ok(Value::Nothing); }
                            _ => {}
                        }
                    }

                    // ===== Nullable instance methods =====
                    if type_name == "Nullable" {
                        match method_name.as_str() {
                            "getvalueordefault" => {
                                let has = obj_ref.borrow().fields.get("hasvalue").and_then(|v| v.as_bool().ok()).unwrap_or(false);
                                if has {
                                    return Ok(obj_ref.borrow().fields.get("value").cloned().unwrap_or(Value::Nothing));
                                }
                                let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                                let arg_values = arg_values?;
                                return Ok(arg_values.get(0).cloned().unwrap_or(Value::Nothing));
                            }
                            _ => {}
                        }
                    }

                    // General .ToString() for typed objects
                    if method_name == "tostring" {
                        if type_name == "Guid" {
                            let val = obj_ref.borrow().fields.get("__value").cloned().unwrap_or(Value::String(String::new()));
                            return Ok(val);
                        }
                        if type_name == "Stopwatch" {
                            let elapsed = obj_ref.borrow().fields.get("elapsedmilliseconds").cloned().unwrap_or(Value::Long(0));
                            return Ok(Value::String(format!("{}ms", elapsed.as_string())));
                        }
                        // Default: return class name
                        return Ok(Value::String(type_name.to_string()));
                    }
                }
            }
            
            // Handle Dialog ShowDialog method
            if method_name == "showdialog" {
                if let Value::Object(obj_ref) = &obj_val {
                    let has_dialog_type = obj_ref.borrow().fields.contains_key("_dialog_type");
                    if has_dialog_type {
                        return crate::builtins::dialogs::dialog_showdialog(&obj_val);
                    }
                }
            }

            // ===== DATABASE OBJECT METHOD DISPATCH =====
            if let Value::Object(obj_ref) = &obj_val {
                let type_name = obj_ref.borrow().fields.get("__type").and_then(|v| {
                    if let Value::String(s) = v { Some(s.clone()) } else { None }
                }).unwrap_or_default();

                // Only evaluate args for known DB types to avoid double-evaluation
                // when the object is a non-DB type (e.g. Namespace for Console)
                let is_db_type = matches!(type_name.as_str(),
                    "DbConnection" | "DbCommand" | "DbRecordset" | "DbReader" |
                    "DbTransaction" | "DataAdapter" | "DbParameters" | "DataRow" |
                    "DataSet" | "DataTable" | "DbParameter" | "BindingSource" | "DataBindings"
                );

                if is_db_type {
                let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                    .map(|arg| self.evaluate_expr(arg))
                    .collect();
                let arg_values = arg_values?;

                // --- DbConnection methods ---
                if type_name == "DbConnection" {
                    match method_name.as_str() {
                        "open" => {
                            // Get or update connection string
                            let conn_str = if !arg_values.is_empty() {
                                let cs = arg_values[0].as_string();
                                obj_ref.borrow_mut().fields.insert("connectionstring".to_string(), Value::String(cs.clone()));
                                cs
                            } else {
                                obj_ref.borrow().fields.get("connectionstring")
                                    .map(|v| v.as_string()).unwrap_or_default()
                            };
                            let dam = crate::data_access::get_global_dam();
                            let conn_id = dam.lock().unwrap().open_connection(&conn_str)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            obj_ref.borrow_mut().fields.insert("__conn_id".to_string(), Value::Long(conn_id as i64));
                            obj_ref.borrow_mut().fields.insert("state".to_string(), Value::Integer(1));
                            return Ok(Value::Nothing);
                        }
                        "close" | "dispose" => {
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            if conn_id > 0 {
                                let dam = crate::data_access::get_global_dam();
                                let _ = dam.lock().unwrap().close_connection(conn_id);
                            }
                            obj_ref.borrow_mut().fields.insert("__conn_id".to_string(), Value::Long(0));
                            obj_ref.borrow_mut().fields.insert("state".to_string(), Value::Integer(0));
                            return Ok(Value::Nothing);
                        }
                        "execute" => {
                            // ADODB-style: conn.Execute(sql) returns a Recordset
                            let sql = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let result = dam.lock().unwrap().execute(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            match result {
                                crate::data_access::ExecuteResult::Recordset(rs_id) => {
                                    let mut fields = std::collections::HashMap::new();
                                    fields.insert("__type".to_string(), Value::String("DbRecordset".to_string()));
                                    fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                    fields.insert("__conn_id".to_string(), Value::Long(conn_id as i64));
                                    let obj = crate::value::ObjectData { class_name: "DbRecordset".to_string(), fields };
                                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                }
                                crate::data_access::ExecuteResult::RowsAffected(n) => {
                                    return Ok(Value::Long(n));
                                }
                            }
                        }
                        "createcommand" => {
                            // ADO.NET: conn.CreateCommand() returns a new Command
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l) } else { None })
                                .unwrap_or(0);
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("__type".to_string(), Value::String("DbCommand".to_string()));
                            fields.insert("__conn_id".to_string(), Value::Long(conn_id));
                            fields.insert("commandtext".to_string(), Value::String(String::new()));
                            fields.insert("commandtype".to_string(), Value::Integer(1));
                            fields.insert("__parameters".to_string(), Value::Collection(
                                std::rc::Rc::new(std::cell::RefCell::new(crate::collections::ArrayList::new()))
                            ));
                            let obj = crate::value::ObjectData { class_name: "DbCommand".to_string(), fields };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                        }
                        "begintransaction" => {
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let tx_id = dam.lock().unwrap().begin_transaction(conn_id)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("__type".to_string(), Value::String("DbTransaction".to_string()));
                            fields.insert("__conn_id".to_string(), Value::Long(tx_id as i64));
                            let obj = crate::value::ObjectData { class_name: "DbTransaction".to_string(), fields };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                        }
                        // Async wrappers — delegate to sync versions
                        "openasync" => {
                            let conn_str = if !arg_values.is_empty() {
                                let cs = arg_values[0].as_string();
                                obj_ref.borrow_mut().fields.insert("connectionstring".to_string(), Value::String(cs.clone()));
                                cs
                            } else {
                                obj_ref.borrow().fields.get("connectionstring")
                                    .map(|v| v.as_string()).unwrap_or_default()
                            };
                            let dam = crate::data_access::get_global_dam();
                            let conn_id = dam.lock().unwrap().open_connection(&conn_str)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            obj_ref.borrow_mut().fields.insert("__conn_id".to_string(), Value::Long(conn_id as i64));
                            obj_ref.borrow_mut().fields.insert("state".to_string(), Value::Integer(1));
                            return Ok(Value::Nothing);
                        }
                        "closeasync" => {
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            if conn_id > 0 {
                                let dam = crate::data_access::get_global_dam();
                                let _ = dam.lock().unwrap().close_connection(conn_id);
                            }
                            obj_ref.borrow_mut().fields.insert("__conn_id".to_string(), Value::Long(0));
                            obj_ref.borrow_mut().fields.insert("state".to_string(), Value::Integer(0));
                            return Ok(Value::Nothing);
                        }
                        "getschema" => {
                            let collection = arg_values.get(0).map(|v| v.as_string()).unwrap_or("Tables".to_string());
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let rs_id = dam.lock().unwrap().get_schema(conn_id, &collection)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            // Return as a DataTable
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("__type".to_string(), Value::String("DataTable".to_string()));
                            fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                            fields.insert("tablename".to_string(), Value::String(collection));
                            let obj = crate::value::ObjectData { class_name: "DataTable".to_string(), fields };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                        }
                        "changedatabase" => {
                            // Just update the connectionstring / database info
                            let new_db = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                            obj_ref.borrow_mut().fields.insert("database".to_string(), Value::String(new_db));
                            return Ok(Value::Nothing);
                        }
                        _ => {}
                    }
                }

                // --- DbCommand methods ---
                if type_name == "DbCommand" {
                    // Helper: collect parameters from __parameters collection
                    let db_params: Vec<(String, String)> = {
                        let borrow = obj_ref.borrow();
                        if let Some(Value::Collection(coll_rc)) = borrow.fields.get("__parameters") {
                            let coll = coll_rc.borrow();
                            let mut pairs = Vec::new();
                            // Items are stored as "@name=value" strings
                            for item in &coll.items {
                                let s = item.as_string();
                                if let Some(eq) = s.find('=') {
                                    let name = s[..eq].to_string();
                                    let val = s[eq+1..].to_string();
                                    pairs.push((name, val));
                                }
                            }
                            pairs
                        } else { Vec::new() }
                    };

                    match method_name.as_str() {
                        // --- Parameters.Add / AddWithValue ---
                        "addparameter" | "addwithvalue" => {
                            // cmd.Parameters.AddWithValue("@name", value)
                            // We handle this on the command object directly
                            let p_name = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                            let p_val = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                            let entry = format!("{}={}", p_name, p_val);
                            let borrow = obj_ref.borrow();
                            if let Some(Value::Collection(coll_rc)) = borrow.fields.get("__parameters") {
                                coll_rc.borrow_mut().items.push(Value::String(entry));
                            }
                            return Ok(Value::Nothing);
                        }
                        "executereader" => {
                            let raw_sql = obj_ref.borrow().fields.get("commandtext")
                                .map(|v| v.as_string()).unwrap_or_default();
                            let cmd_type = obj_ref.borrow().fields.get("commandtype")
                                .and_then(|v| if let Value::Integer(i) = v { Some(*i) } else { None })
                                .unwrap_or(1);
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            // CommandType: 1=Text, 4=StoredProcedure
                            if cmd_type == 4 {
                                let dam = crate::data_access::get_global_dam();
                                let result = dam.lock().unwrap().execute_stored_proc(conn_id, &raw_sql, &db_params)
                                    .map_err(|e| RuntimeError::Custom(e))?;
                                match result {
                                    crate::data_access::ExecuteResult::Recordset(rs_id) => {
                                        let mut fields = std::collections::HashMap::new();
                                        fields.insert("__type".to_string(), Value::String("DbReader".to_string()));
                                        fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                        fields.insert("__started".to_string(), Value::Boolean(false));
                                        let obj = crate::value::ObjectData { class_name: "DbReader".to_string(), fields };
                                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                    }
                                    crate::data_access::ExecuteResult::RowsAffected(n) => return Ok(Value::Integer(n as i32)),
                                }
                            }
                            let sql = if db_params.is_empty() { raw_sql } else {
                                crate::data_access::DataAccessManager::substitute_params(&raw_sql, &db_params)
                            };
                            let dam = crate::data_access::get_global_dam();
                            let rs_id = dam.lock().unwrap().execute_reader(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            // Return a DbReader (ADO.NET SqlDataReader equivalent)
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("__type".to_string(), Value::String("DbReader".to_string()));
                            fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                            fields.insert("__started".to_string(), Value::Boolean(false));
                            let obj = crate::value::ObjectData { class_name: "DbReader".to_string(), fields };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                        }
                        "executenonquery" => {
                            let raw_sql = obj_ref.borrow().fields.get("commandtext")
                                .map(|v| v.as_string()).unwrap_or_default();
                            let sql = if db_params.is_empty() { raw_sql } else {
                                crate::data_access::DataAccessManager::substitute_params(&raw_sql, &db_params)
                            };
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let affected = dam.lock().unwrap().execute_non_query(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            return Ok(Value::Integer(affected as i32));
                        }
                        "executescalar" => {
                            let raw_sql = obj_ref.borrow().fields.get("commandtext")
                                .map(|v| v.as_string()).unwrap_or_default();
                            let sql = if db_params.is_empty() { raw_sql } else {
                                crate::data_access::DataAccessManager::substitute_params(&raw_sql, &db_params)
                            };
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let result = dam.lock().unwrap().execute_scalar(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            // Try to parse as number, else return string
                            if let Ok(i) = result.parse::<i32>() {
                                return Ok(Value::Integer(i));
                            }
                            if let Ok(f) = result.parse::<f64>() {
                                return Ok(Value::Double(f));
                            }
                            return Ok(Value::String(result));
                        }
                        "execute" => {
                            // ADODB-style: cmd.Execute — auto-detect SELECT vs DML
                            let raw_sql = obj_ref.borrow().fields.get("commandtext")
                                .map(|v| v.as_string()).unwrap_or_default();
                            let cmd_type = obj_ref.borrow().fields.get("commandtype")
                                .and_then(|v| if let Value::Integer(i) = v { Some(*i) } else { None })
                                .unwrap_or(1);
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            if cmd_type == 4 {
                                // Stored procedure
                                let result = dam.lock().unwrap().execute_stored_proc(conn_id, &raw_sql, &db_params)
                                    .map_err(|e| RuntimeError::Custom(e))?;
                                match result {
                                    crate::data_access::ExecuteResult::Recordset(rs_id) => {
                                        let mut fields = std::collections::HashMap::new();
                                        fields.insert("__type".to_string(), Value::String("DbRecordset".to_string()));
                                        fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                        fields.insert("__conn_id".to_string(), Value::Long(conn_id as i64));
                                        let obj = crate::value::ObjectData { class_name: "DbRecordset".to_string(), fields };
                                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                    }
                                    crate::data_access::ExecuteResult::RowsAffected(n) => return Ok(Value::Long(n)),
                                }
                            }
                            let sql = if db_params.is_empty() { raw_sql } else {
                                crate::data_access::DataAccessManager::substitute_params(&raw_sql, &db_params)
                            };
                            let result = dam.lock().unwrap().execute(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            match result {
                                crate::data_access::ExecuteResult::Recordset(rs_id) => {
                                    let mut fields = std::collections::HashMap::new();
                                    fields.insert("__type".to_string(), Value::String("DbRecordset".to_string()));
                                    fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                    fields.insert("__conn_id".to_string(), Value::Long(conn_id as i64));
                                    let obj = crate::value::ObjectData { class_name: "DbRecordset".to_string(), fields };
                                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                }
                                crate::data_access::ExecuteResult::RowsAffected(n) => return Ok(Value::Long(n)),
                            }
                        }
                        // Async wrappers — delegate to sync versions
                        "executereaderasync" => {
                            let raw_sql = obj_ref.borrow().fields.get("commandtext")
                                .map(|v| v.as_string()).unwrap_or_default();
                            let cmd_type = obj_ref.borrow().fields.get("commandtype")
                                .and_then(|v| if let Value::Integer(i) = v { Some(*i) } else { None })
                                .unwrap_or(1);
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let sql = if cmd_type == 4 {
                                // StoredProcedure type
                                let dam = crate::data_access::get_global_dam();
                                let result = dam.lock().unwrap().execute_stored_proc(conn_id, &raw_sql, &db_params)
                                    .map_err(|e| RuntimeError::Custom(e))?;
                                match result {
                                    crate::data_access::ExecuteResult::Recordset(rs_id) => {
                                        let mut fields = std::collections::HashMap::new();
                                        fields.insert("__type".to_string(), Value::String("DbReader".to_string()));
                                        fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                        fields.insert("__started".to_string(), Value::Boolean(false));
                                        let obj = crate::value::ObjectData { class_name: "DbReader".to_string(), fields };
                                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                    }
                                    crate::data_access::ExecuteResult::RowsAffected(n) => return Ok(Value::Integer(n as i32)),
                                }
                            } else if db_params.is_empty() { raw_sql } else {
                                crate::data_access::DataAccessManager::substitute_params(&raw_sql, &db_params)
                            };
                            let dam = crate::data_access::get_global_dam();
                            let rs_id = dam.lock().unwrap().execute_reader(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("__type".to_string(), Value::String("DbReader".to_string()));
                            fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                            fields.insert("__started".to_string(), Value::Boolean(false));
                            let obj = crate::value::ObjectData { class_name: "DbReader".to_string(), fields };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                        }
                        "executenonqueryasync" => {
                            let raw_sql = obj_ref.borrow().fields.get("commandtext")
                                .map(|v| v.as_string()).unwrap_or_default();
                            let sql = if db_params.is_empty() { raw_sql } else {
                                crate::data_access::DataAccessManager::substitute_params(&raw_sql, &db_params)
                            };
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let affected = dam.lock().unwrap().execute_non_query(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            return Ok(Value::Integer(affected as i32));
                        }
                        "executescalarasync" => {
                            let raw_sql = obj_ref.borrow().fields.get("commandtext")
                                .map(|v| v.as_string()).unwrap_or_default();
                            let sql = if db_params.is_empty() { raw_sql } else {
                                crate::data_access::DataAccessManager::substitute_params(&raw_sql, &db_params)
                            };
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let result = dam.lock().unwrap().execute_scalar(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            if let Ok(i) = result.parse::<i32>() { return Ok(Value::Integer(i)); }
                            if let Ok(f) = result.parse::<f64>() { return Ok(Value::Double(f)); }
                            return Ok(Value::String(result));
                        }
                        // Stored procedure execution
                        "executeproc" | "executestoredproc" => {
                            let proc_name = obj_ref.borrow().fields.get("commandtext")
                                .map(|v| v.as_string()).unwrap_or_default();
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let result = dam.lock().unwrap().execute_stored_proc(conn_id, &proc_name, &db_params)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            match result {
                                crate::data_access::ExecuteResult::Recordset(rs_id) => {
                                    let mut fields = std::collections::HashMap::new();
                                    fields.insert("__type".to_string(), Value::String("DbReader".to_string()));
                                    fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                    fields.insert("__started".to_string(), Value::Boolean(false));
                                    let obj = crate::value::ObjectData { class_name: "DbReader".to_string(), fields };
                                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                }
                                crate::data_access::ExecuteResult::RowsAffected(n) => return Ok(Value::Integer(n as i32)),
                            }
                        }
                        "cancel" | "dispose" => {
                            return Ok(Value::Nothing);
                        }
                        "prepare" => {
                            // No-op for our implementation
                            return Ok(Value::Nothing);
                        }
                        "createparameter" => {
                            // ADODB-style: cmd.CreateParameter(name, type, direction, size, value)
                            let p_name = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                            let p_type = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(200);
                            let p_dir = arg_values.get(2).and_then(|v| v.as_integer().ok()).unwrap_or(1);
                            let p_size = arg_values.get(3).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                            let p_val = arg_values.get(4).cloned().unwrap_or(Value::Nothing);
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("__type".to_string(), Value::String("DbParameter".to_string()));
                            fields.insert("name".to_string(), Value::String(p_name));
                            fields.insert("value".to_string(), p_val);
                            fields.insert("direction".to_string(), Value::Integer(p_dir));
                            fields.insert("type".to_string(), Value::Integer(p_type));
                            fields.insert("size".to_string(), Value::Integer(p_size));
                            let obj = crate::value::ObjectData { class_name: "DbParameter".to_string(), fields };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                        }
                        _ => {}
                    }
                }

                // --- DbRecordset methods (ADODB Recordset API) ---
                if type_name == "DbRecordset" {
                    match method_name.as_str() {
                        "open" => {
                            // rs.Open sql, conn
                            let sql = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                            let conn_id = if let Some(conn_val) = arg_values.get(1) {
                                if let Value::Object(cr) = conn_val {
                                    cr.borrow().fields.get("__conn_id")
                                        .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                        .unwrap_or(0)
                                } else { 0 }
                            } else {
                                obj_ref.borrow().fields.get("__conn_id")
                                    .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                    .unwrap_or(0)
                            };
                            let dam = crate::data_access::get_global_dam();
                            let rs_id = dam.lock().unwrap().execute_reader(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            obj_ref.borrow_mut().fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                            obj_ref.borrow_mut().fields.insert("__conn_id".to_string(), Value::Long(conn_id as i64));
                            return Ok(Value::Nothing);
                        }
                        "movenext" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            if let Some(rs) = dam.lock().unwrap().recordsets.get_mut(&rs_id) {
                                rs.move_next();
                            }
                            return Ok(Value::Nothing);
                        }
                        "movefirst" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            if let Some(rs) = dam.lock().unwrap().recordsets.get_mut(&rs_id) {
                                rs.move_first();
                            }
                            return Ok(Value::Nothing);
                        }
                        "movelast" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            if let Some(rs) = dam.lock().unwrap().recordsets.get_mut(&rs_id) {
                                rs.move_last();
                            }
                            return Ok(Value::Nothing);
                        }
                        "moveprevious" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            if let Some(rs) = dam.lock().unwrap().recordsets.get_mut(&rs_id) {
                                rs.move_previous();
                            }
                            return Ok(Value::Nothing);
                        }
                        "close" | "dispose" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            dam.lock().unwrap().close_recordset(rs_id);
                            return Ok(Value::Nothing);
                        }
                        "fields" => {
                            // rs.Fields("name") or rs.Fields(0) — return a field-accessor object
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                if let Some(row) = rs.current_row() {
                                    let (val, field_name) = match &arg_values[0] {
                                        Value::Integer(i) => {
                                            let idx = *i as usize;
                                            let v = row.get_by_index(idx).unwrap_or("NULL").to_string();
                                            let n = rs.columns.get(idx).cloned().unwrap_or_default();
                                            (v, n)
                                        }
                                        _ => {
                                            let name = arg_values[0].as_string();
                                            let v = row.get_by_name(&name).unwrap_or("NULL").to_string();
                                            (v, name)
                                        }
                                    };
                                    // Return a Field object with .Value property
                                    let mut fld = std::collections::HashMap::new();
                                    fld.insert("__type".to_string(), Value::String("DbField".to_string()));
                                    fld.insert("value".to_string(), Value::String(val.clone()));
                                    fld.insert("name".to_string(), Value::String(field_name));
                                    let obj = crate::value::ObjectData { class_name: "DbField".to_string(), fields: fld };
                                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        _ => {}
                    }
                }

                // --- DbReader methods (ADO.NET SqlDataReader API) ---
                if type_name == "DbReader" {
                    match method_name.as_str() {
                        "read" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let started = obj_ref.borrow().fields.get("__started")
                                .and_then(|v| if let Value::Boolean(b) = v { Some(*b) } else { None })
                                .unwrap_or(false);
                            let dam = crate::data_access::get_global_dam();
                            let has_row = if let Some(rs) = dam.lock().unwrap().recordsets.get_mut(&rs_id) {
                                if !started {
                                    // First Read() — check if any rows exist
                                    !rs.eof()
                                } else {
                                    rs.move_next();
                                    !rs.eof()
                                }
                            } else { false };
                            obj_ref.borrow_mut().fields.insert("__started".to_string(), Value::Boolean(true));
                            return Ok(Value::Boolean(has_row));
                        }
                        "getstring" | "getvalue" | "item" => {
                            // reader.GetString(ordinal) or reader("name")
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                if let Some(row) = rs.current_row() {
                                    let val = match &arg_values[0] {
                                        Value::Integer(i) => row.get_by_index(*i as usize)
                                            .unwrap_or("NULL").to_string(),
                                        Value::String(s) => row.get_by_name(s)
                                            .unwrap_or("NULL").to_string(),
                                        _ => arg_values[0].as_string(),
                                    };
                                    return Ok(Value::String(val));
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        "getint32" | "getint16" | "getint64" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                if let Some(row) = rs.current_row() {
                                    let idx = arg_values[0].as_integer().unwrap_or(0) as usize;
                                    let val_str = row.get_by_index(idx).unwrap_or("0");
                                    let val = val_str.parse::<i32>().unwrap_or(0);
                                    return Ok(Value::Integer(val));
                                }
                            }
                            return Ok(Value::Integer(0));
                        }
                        "getdouble" | "getfloat" | "getdecimal" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                if let Some(row) = rs.current_row() {
                                    let idx = arg_values[0].as_integer().unwrap_or(0) as usize;
                                    let val_str = row.get_by_index(idx).unwrap_or("0");
                                    let val = val_str.parse::<f64>().unwrap_or(0.0);
                                    return Ok(Value::Double(val));
                                }
                            }
                            return Ok(Value::Double(0.0));
                        }
                        "getboolean" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                if let Some(row) = rs.current_row() {
                                    let idx = arg_values[0].as_integer().unwrap_or(0) as usize;
                                    let val_str = row.get_by_index(idx).unwrap_or("false");
                                    let val = val_str == "1" || val_str.eq_ignore_ascii_case("true");
                                    return Ok(Value::Boolean(val));
                                }
                            }
                            return Ok(Value::Boolean(false));
                        }
                        "getordinal" => {
                            let col_name = arg_values[0].as_string();
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                let lower = col_name.to_lowercase();
                                for (i, c) in rs.columns.iter().enumerate() {
                                    if c.to_lowercase() == lower {
                                        return Ok(Value::Integer(i as i32));
                                    }
                                }
                            }
                            return Ok(Value::Integer(-1));
                        }
                        "close" | "dispose" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            dam.lock().unwrap().close_recordset(rs_id);
                            return Ok(Value::Nothing);
                        }
                        "isdbnull" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                if let Some(row) = rs.current_row() {
                                    let idx = arg_values[0].as_integer().unwrap_or(0) as usize;
                                    let val_str = row.get_by_index(idx).unwrap_or("NULL");
                                    return Ok(Value::Boolean(val_str == "NULL"));
                                }
                            }
                            return Ok(Value::Boolean(true));
                        }
                        "getname" => {
                            // reader.GetName(ordinal) — return column name
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                let idx = arg_values[0].as_integer().unwrap_or(0) as usize;
                                if idx < rs.columns.len() {
                                    return Ok(Value::String(rs.columns[idx].clone()));
                                }
                            }
                            return Ok(Value::String(String::new()));
                        }
                        "nextresult" => {
                            // Advance to the next result set (for multi-SELECT)
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let has_next = dam.lock().unwrap().next_result(rs_id);
                            if has_next {
                                // Reset the started flag for the new result set
                                obj_ref.borrow_mut().fields.insert("__started".to_string(), Value::Boolean(false));
                            }
                            return Ok(Value::Boolean(has_next));
                        }
                        "readasync" => {
                            // Async wrapper for Read — same behavior
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let started = obj_ref.borrow().fields.get("__started")
                                .and_then(|v| if let Value::Boolean(b) = v { Some(*b) } else { None })
                                .unwrap_or(false);
                            let dam = crate::data_access::get_global_dam();
                            let has_row = if let Some(rs) = dam.lock().unwrap().recordsets.get_mut(&rs_id) {
                                if !started { !rs.eof() } else { rs.move_next(); !rs.eof() }
                            } else { false };
                            obj_ref.borrow_mut().fields.insert("__started".to_string(), Value::Boolean(true));
                            return Ok(Value::Boolean(has_row));
                        }
                        "getschematable" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let schema_id = dam.lock().unwrap().get_schema_table(rs_id)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("__type".to_string(), Value::String("DataTable".to_string()));
                            fields.insert("__rs_id".to_string(), Value::Long(schema_id as i64));
                            fields.insert("tablename".to_string(), Value::String("SchemaTable".to_string()));
                            let obj = crate::value::ObjectData { class_name: "DataTable".to_string(), fields };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                        }
                        "getdatetime" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                if let Some(row) = rs.current_row() {
                                    let idx = arg_values[0].as_integer().unwrap_or(0) as usize;
                                    let val_str = row.get_by_index(idx).unwrap_or("");
                                    return Ok(Value::String(val_str.to_string()));
                                }
                            }
                            return Ok(Value::String(String::new()));
                        }
                        "getguid" => {
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                if let Some(row) = rs.current_row() {
                                    let idx = arg_values[0].as_integer().unwrap_or(0) as usize;
                                    let val_str = row.get_by_index(idx).unwrap_or("");
                                    return Ok(Value::String(val_str.to_string()));
                                }
                            }
                            return Ok(Value::String(String::new()));
                        }
                        "getfieldtype" => {
                            // Returns type name for a column — we always use String
                            return Ok(Value::String("String".to_string()));
                        }
                        _ => {}
                    }
                }

                // --- DbTransaction methods ---
                if type_name == "DbTransaction" {
                    match method_name.as_str() {
                        "commit" => {
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            dam.lock().unwrap().commit(conn_id)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            return Ok(Value::Nothing);
                        }
                        "rollback" => {
                            let conn_id = obj_ref.borrow().fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            dam.lock().unwrap().rollback(conn_id)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            return Ok(Value::Nothing);
                        }
                        "dispose" => return Ok(Value::Nothing),
                        _ => {}
                    }
                }

                // --- DataAdapter methods ---
                if type_name == "DataAdapter" {
                    match method_name.as_str() {
                        "fill" => {
                            // adapter.Fill(dataTable) or adapter.Fill(dataSet)
                            let da_borrow = obj_ref.borrow();
                            let sql = da_borrow.fields.get("selectcommandtext")
                                .or_else(|| da_borrow.fields.get("selectcommand"))
                                .map(|v| v.as_string()).unwrap_or_default();
                            let conn_id = da_borrow.fields.get("__conn_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            drop(da_borrow);
                            let dam = crate::data_access::get_global_dam();
                            let rs_id = dam.lock().unwrap().execute_reader(conn_id, &sql)
                                .map_err(|e| RuntimeError::Custom(e))?;
                            let row_count = dam.lock().unwrap().recordsets.get(&rs_id)
                                .map(|rs| rs.record_count()).unwrap_or(0);
                            // Check if arg is DataTable or DataSet
                            if let Some(dt_val) = arg_values.first() {
                                if let Value::Object(dt_ref) = dt_val {
                                    let dt_type = dt_ref.borrow().fields.get("__type")
                                        .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                                        .unwrap_or_default();
                                    if dt_type == "DataSet" {
                                        // Create a DataTable and add it to the DataSet's tables
                                        let mut dt_fields = std::collections::HashMap::new();
                                        dt_fields.insert("__type".to_string(), Value::String("DataTable".to_string()));
                                        dt_fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                        dt_fields.insert("tablename".to_string(), Value::String("Table".to_string()));
                                        let new_dt = Value::Object(std::rc::Rc::new(std::cell::RefCell::new(
                                            crate::value::ObjectData { class_name: "DataTable".to_string(), fields: dt_fields }
                                        )));
                                        let mut ds_borrow = dt_ref.borrow_mut();
                                        if let Some(Value::Array(tables)) = ds_borrow.fields.get_mut("__tables") {
                                            tables.push(new_dt);
                                        }
                                    } else {
                                        // DataTable
                                        dt_ref.borrow_mut().fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                    }
                                }
                            }
                            return Ok(Value::Integer(row_count));
                        }
                        "dispose" => return Ok(Value::Nothing),
                        _ => {}
                    }
                }

                // --- DbParameters collection (for cmd.Parameters.AddWithValue) ---
                if type_name == "DbParameters" {
                    match method_name.as_str() {
                        "addwithvalue" | "add" => {
                            let p_name = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                            let p_val = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                            // Store back to parent command's __parameters
                            let parent_cmd = obj_ref.borrow().fields.get("__parent_cmd").cloned();
                            if let Some(Value::Object(cmd_ref)) = parent_cmd {
                                let cmd_borrow = cmd_ref.borrow();
                                if let Some(Value::Collection(coll_rc)) = cmd_borrow.fields.get("__parameters") {
                                    let entry = format!("{}={}", p_name, p_val);
                                    coll_rc.borrow_mut().items.push(Value::String(entry));
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        "append" => {
                            // ADODB-style: cmd.Parameters.Append(param)
                            // param is a DbParameter object with name and value fields
                            if let Some(param_val) = arg_values.first() {
                                if let Value::Object(param_ref) = param_val {
                                    let p_name = param_ref.borrow().fields.get("name")
                                        .map(|v| v.as_string()).unwrap_or_default();
                                    let p_val = param_ref.borrow().fields.get("value")
                                        .map(|v| v.as_string()).unwrap_or_default();
                                    let parent_cmd = obj_ref.borrow().fields.get("__parent_cmd").cloned();
                                    if let Some(Value::Object(cmd_ref)) = parent_cmd {
                                        let cmd_borrow = cmd_ref.borrow();
                                        if let Some(Value::Collection(coll_rc)) = cmd_borrow.fields.get("__parameters") {
                                            let name_with_at = if p_name.starts_with('@') { p_name } else { format!("@{}", p_name) };
                                            let entry = format!("{}={}", name_with_at, p_val);
                                            coll_rc.borrow_mut().items.push(Value::String(entry));
                                        }
                                    }
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        "clear" => {
                            let parent_cmd = obj_ref.borrow().fields.get("__parent_cmd").cloned();
                            if let Some(Value::Object(cmd_ref)) = parent_cmd {
                                let cmd_borrow = cmd_ref.borrow();
                                if let Some(Value::Collection(coll_rc)) = cmd_borrow.fields.get("__parameters") {
                                    coll_rc.borrow_mut().items.clear();
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        "count" => {
                            let parent_cmd = obj_ref.borrow().fields.get("__parent_cmd").cloned();
                            if let Some(Value::Object(cmd_ref)) = parent_cmd {
                                let cmd_borrow = cmd_ref.borrow();
                                if let Some(Value::Collection(coll_rc)) = cmd_borrow.fields.get("__parameters") {
                                    return Ok(Value::Integer(coll_rc.borrow().items.len() as i32));
                                }
                            }
                            return Ok(Value::Integer(0));
                        }
                        "item" => {
                            // Parameters.Item(index) or Parameters.Item("name")
                            let parent_cmd = obj_ref.borrow().fields.get("__parent_cmd").cloned();
                            if let Some(Value::Object(cmd_ref)) = parent_cmd {
                                let cmd_borrow = cmd_ref.borrow();
                                if let Some(Value::Collection(coll_rc)) = cmd_borrow.fields.get("__parameters") {
                                    let coll = coll_rc.borrow();
                                    match &arg_values[0] {
                                        Value::Integer(idx) => {
                                            if let Some(item) = coll.items.get(*idx as usize) {
                                                return Ok(item.clone());
                                            }
                                        }
                                        _ => {
                                            let name = arg_values[0].as_string();
                                            for item in &coll.items {
                                                let s = item.as_string();
                                                if s.starts_with(&format!("{}=", name)) {
                                                    return Ok(item.clone());
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        "removeat" | "remove" => {
                            let parent_cmd = obj_ref.borrow().fields.get("__parent_cmd").cloned();
                            if let Some(Value::Object(cmd_ref)) = parent_cmd {
                                let cmd_borrow = cmd_ref.borrow();
                                if let Some(Value::Collection(coll_rc)) = cmd_borrow.fields.get("__parameters") {
                                    let mut coll = coll_rc.borrow_mut();
                                    match &arg_values[0] {
                                        Value::Integer(idx) => {
                                            let i = *idx as usize;
                                            if i < coll.items.len() { coll.items.remove(i); }
                                        }
                                        _ => {
                                            let name = arg_values[0].as_string();
                                            coll.items.retain(|item| !item.as_string().starts_with(&format!("{}=", name)));
                                        }
                                    }
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        _ => {}
                    }
                }

                // --- DataRow methods ---
                if type_name == "DataRow" {
                    match method_name.as_str() {
                        "item" | "" => {
                            // row.Item("colname") or row.Item(index) or row("colname") default indexer
                            match &arg_values[0] {
                                Value::Integer(i) => {
                                    let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                        .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                        .unwrap_or(0);
                                    let row_idx = obj_ref.borrow().fields.get("__row_index")
                                        .and_then(|v| if let Value::Integer(i) = v { Some(*i as usize) } else { None })
                                        .unwrap_or(0);
                                    let dam = crate::data_access::get_global_dam();
                                    let dam_lock = dam.lock().unwrap();
                                    if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                        if let Some(row) = rs.rows.get(row_idx) {
                                            let v = row.get_by_index(*i as usize).unwrap_or("NULL").to_string();
                                            return Ok(Value::String(v));
                                        }
                                    }
                                    return Ok(Value::String("NULL".to_string()));
                                }
                                _ => {
                                    let col = arg_values[0].as_string().to_lowercase();
                                    let val = obj_ref.borrow().fields.get(&col)
                                        .cloned().unwrap_or(Value::String("NULL".to_string()));
                                    return Ok(val);
                                }
                            }
                        }
                        "isnull" => {
                            let col = arg_values[0].as_string().to_lowercase();
                            let val = obj_ref.borrow().fields.get(&col)
                                .map(|v| v.as_string() == "NULL").unwrap_or(true);
                            return Ok(Value::Boolean(val));
                        }
                        "tostring" => {
                            // Return a string representation of the row
                            let obj_data = obj_ref.borrow();
                            let mut parts = Vec::new();
                            for (k, v) in &obj_data.fields {
                                if !k.starts_with("__") {
                                    parts.push(format!("{}={}", k, v.as_string()));
                                }
                            }
                            return Ok(Value::String(parts.join(", ")));
                        }
                        _ => {}
                    }
                }

                // --- DataTable methods ---
                if type_name == "DataTable" {
                    match method_name.as_str() {
                        "select" => {
                            // dt.Select() — returns all rows as an array
                            // dt.Select(filter) — basic filtering (not implemented, returns all)
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                let mut row_objects = Vec::new();
                                for (i, db_row) in rs.rows.iter().enumerate() {
                                    let mut flds = std::collections::HashMap::new();
                                    flds.insert("__type".to_string(), Value::String("DataRow".to_string()));
                                    flds.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                                    flds.insert("__row_index".to_string(), Value::Integer(i as i32));
                                    for (ci, col) in db_row.columns.iter().enumerate() {
                                        let v = db_row.values.get(ci).cloned().unwrap_or_default();
                                        flds.insert(col.to_lowercase(), Value::String(v));
                                    }
                                    let obj = crate::value::ObjectData { class_name: "DataRow".to_string(), fields: flds };
                                    row_objects.push(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                                }
                                return Ok(Value::Array(row_objects));
                            }
                            return Ok(Value::Array(Vec::new()));
                        }
                        "newrow" => {
                            // Create a new empty DataRow for this table
                            let rs_id = obj_ref.borrow().fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            let mut flds = std::collections::HashMap::new();
                            flds.insert("__type".to_string(), Value::String("DataRow".to_string()));
                            flds.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                            flds.insert("__row_index".to_string(), Value::Integer(-1));
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                for col in &rs.columns {
                                    flds.insert(col.to_lowercase(), Value::String(String::new()));
                                }
                            }
                            let obj = crate::value::ObjectData { class_name: "DataRow".to_string(), fields: flds };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                        }
                        "clear" | "dispose" => {
                            return Ok(Value::Nothing);
                        }
                        "copy" | "clone" => {
                            // Return a shallow copy
                            let obj_data = obj_ref.borrow();
                            let new_fields = obj_data.fields.clone();
                            let obj = crate::value::ObjectData { class_name: "DataTable".to_string(), fields: new_fields };
                            return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                        }
                        _ => {}
                    }
                }

                // --- DataSet methods ---
                if type_name == "DataSet" {
                    match method_name.as_str() {
                        "tables" => {
                            // ds.Tables(index) or ds.Tables("name")
                            let tables = obj_ref.borrow().fields.get("__tables")
                                .cloned().unwrap_or(Value::Array(Vec::new()));
                            if let Value::Array(arr) = &tables {
                                match &arg_values[0] {
                                    Value::Integer(idx) => {
                                        if let Some(t) = arr.get(*idx as usize) {
                                            return Ok(t.clone());
                                        }
                                    }
                                    _ => {
                                        let name = arg_values[0].as_string().to_lowercase();
                                        for t in arr {
                                            if let Value::Object(dt_ref) = t {
                                                let tn = dt_ref.borrow().fields.get("tablename")
                                                    .map(|v| v.as_string().to_lowercase()).unwrap_or_default();
                                                if tn == name {
                                                    return Ok(t.clone());
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        "dispose" | "clear" => return Ok(Value::Nothing),
                        _ => {}
                    }
                }

                // --- DbParameter methods ---
                if type_name == "DbParameter" {
                    // Allow setting value on parameter objects
                    return Ok(Value::Nothing);
                }

                // --- BindingSource methods ---
                if type_name == "BindingSource" {
                    let bs_name = obj_ref.borrow().fields.get("name")
                        .map(|v| v.as_string()).unwrap_or_default();
                    match method_name.as_str() {
                        "movenext" => {
                            let bs_val = Value::Object(obj_ref.clone());
                            let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                            let count = self.binding_source_row_count_filtered(&bs_val);
                            let pos = obj_ref.borrow().fields.get("position")
                                .and_then(|v| if let Value::Integer(i) = v { Some(*i) } else { None })
                                .unwrap_or(0);
                            let new_pos = if pos < count - 1 { pos + 1 } else { pos };
                            obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Integer(new_pos));
                            self.side_effects.push_back(crate::RuntimeSideEffect::BindingPositionChanged {
                                binding_source_name: bs_name.clone(),
                                position: new_pos,
                                count,
                            });
                            self.refresh_bindings_filtered(&obj_ref, &ds, new_pos);
                            return Ok(Value::Nothing);
                        }
                        "moveprevious" => {
                            let bs_val = Value::Object(obj_ref.clone());
                            let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                            let count = self.binding_source_row_count_filtered(&bs_val);
                            let pos = obj_ref.borrow().fields.get("position")
                                .and_then(|v| if let Value::Integer(i) = v { Some(*i) } else { None })
                                .unwrap_or(0);
                            let new_pos = if pos > 0 { pos - 1 } else { 0 };
                            obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Integer(new_pos));
                            self.side_effects.push_back(crate::RuntimeSideEffect::BindingPositionChanged {
                                binding_source_name: bs_name.clone(),
                                position: new_pos,
                                count,
                            });
                            self.refresh_bindings_filtered(&obj_ref, &ds, new_pos);
                            return Ok(Value::Nothing);
                        }
                        "movefirst" => {
                            let bs_val = Value::Object(obj_ref.clone());
                            let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                            let count = self.binding_source_row_count_filtered(&bs_val);
                            obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Integer(0));
                            self.side_effects.push_back(crate::RuntimeSideEffect::BindingPositionChanged {
                                binding_source_name: bs_name.clone(),
                                position: 0,
                                count,
                            });
                            self.refresh_bindings_filtered(&obj_ref, &ds, 0);
                            return Ok(Value::Nothing);
                        }
                        "movelast" => {
                            let bs_val = Value::Object(obj_ref.clone());
                            let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                            let count = self.binding_source_row_count_filtered(&bs_val);
                            let last = if count > 0 { count - 1 } else { 0 };
                            obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Integer(last));
                            self.side_effects.push_back(crate::RuntimeSideEffect::BindingPositionChanged {
                                binding_source_name: bs_name.clone(),
                                position: last,
                                count,
                            });
                            self.refresh_bindings_filtered(&obj_ref, &ds, last);
                            return Ok(Value::Nothing);
                        }
                        "removecurrent" => {
                            return Ok(Value::Nothing);
                        }
                        "addnew" => {
                            return Ok(Value::Nothing);
                        }
                        "endedit" | "cancelcurrentedit" | "resetbindings" => {
                            return Ok(Value::Nothing);
                        }
                        "dispose" => return Ok(Value::Nothing),
                        _ => {}
                    }
                }

                // --- DataBindings methods ---
                if type_name == "DataBindings" {
                    match method_name.as_str() {
                        "add" => {
                            // control.DataBindings.Add(propertyName, dataSource, dataMember)
                            let prop_name = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                            let data_source = arg_values.get(1).cloned().unwrap_or(Value::Nothing);
                            let data_member = arg_values.get(2).map(|v| v.as_string()).unwrap_or_default();

                            // Get the parent control name
                            let parent_name = obj_ref.borrow().fields.get("__parent_name")
                                .map(|v| v.as_string())
                                .or_else(|| {
                                    // Fallback to parent object's name field
                                    if let Some(Value::Object(p)) = obj_ref.borrow().fields.get("__parent") {
                                        p.borrow().fields.get("name").map(|v| v.as_string())
                                    } else { None }
                                })
                                .unwrap_or_else(|| "UnknownControl".to_string());

                            // Store binding info in environment for the control
                            let binding_key = format!("__binding_{}_{}", parent_name, prop_name);
                            let mut binding_fields = std::collections::HashMap::new();
                            binding_fields.insert("property".to_string(), Value::String(prop_name.clone()));
                            binding_fields.insert("datasource".to_string(), data_source.clone());
                            binding_fields.insert("datamember".to_string(), Value::String(data_member.clone()));
                            binding_fields.insert("controlname".to_string(), Value::String(parent_name.clone()));
                            let binding_obj = crate::value::ObjectData { class_name: "Binding".to_string(), fields: binding_fields };
                            self.env.define_global(&binding_key, Value::Object(std::rc::Rc::new(std::cell::RefCell::new(binding_obj))));

                            // Resolve data_source to object if it's a string proxy
                            let ds_obj_val = match &data_source {
                                Value::String(proxy_name) => self.resolve_control_as_sender(proxy_name),
                                _ => data_source.clone(),
                            };

                            // Try to immediately sync — get current value from BindingSource
                            if let Value::Object(ds_ref) = ds_obj_val {
                                let ds_type = ds_ref.borrow().fields.get("__type")
                                    .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                                    .unwrap_or_default();
                                if ds_type == "BindingSource" {
                                    // 1. Register binding on the BindingSource object
                                    {
                                        let mut ds_borrow = ds_ref.borrow_mut();
                                        if let Some(Value::Array(bindings)) = ds_borrow.fields.get_mut("__bindings") {
                                            let binding_entry = format!("{}|{}|{}", parent_name, prop_name, data_member);
                                            bindings.push(Value::String(binding_entry));
                                        }
                                    }

                                    // 2. Perform initial sync
                                    let ds_val = ds_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                                    let dm = ds_ref.borrow().fields.get("datamember")
                                        .map(|v| v.as_string()).unwrap_or_default();
                                    
                                    Self::inject_select_from_data_member(&ds_val, &dm);
                                    let pos = ds_ref.borrow().fields.get("position")
                                        .and_then(|v| if let Value::Integer(i) = v { Some(*i) } else { None })
                                        .unwrap_or(0);
                                    let row = self.binding_source_get_row(&ds_val, pos);
                                    if let Value::Object(row_ref) = &row {
                                        let member_lower = data_member.to_lowercase();
                                        let cell_val = row_ref.borrow().fields.get(&member_lower)
                                            .cloned().unwrap_or(Value::String(String::new()));
                                        
                                        self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                                            object: parent_name,
                                            property: prop_name,
                                            value: cell_val,
                                        });
                                    }
                                    // Push initial position update for any linked BindingNavigators
                                    let bs_name = ds_ref.borrow().fields.get("name")
                                        .map(|v| v.as_string()).unwrap_or_default();
                                    if !bs_name.is_empty() {
                                        let count = self.binding_source_row_count(&Value::Object(ds_ref.clone()));
                                        self.side_effects.push_back(crate::RuntimeSideEffect::BindingPositionChanged {
                                            binding_source_name: bs_name,
                                            position: pos,
                                            count,
                                        });
                                    }
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        "clear" => return Ok(Value::Nothing),
                        "remove" | "removeat" => return Ok(Value::Nothing),
                        _ => {}
                    }
                }

            } // end if is_db_type
            }
            
            // Array properties accessed as method calls
            if let Value::Array(arr) = &obj_val {
                match method_name.as_str() {
                    "length" | "count" if args.is_empty() => {
                        return Ok(Value::Integer(arr.len() as i32));
                    }
                    _ => {}
                }
            }

            if let Value::Collection(col_rc) = &obj_val {
                 match method_name.as_str() {
                    "add" => {
                        // VB.NET Collection.Add(item [, key [, before | after]])
                        let val = self.evaluate_expr(&args[0])?;
                        if args.len() >= 2 {
                            let key_val = self.evaluate_expr(&args[1])?;
                            let key_str = key_val.as_string();
                            if args.len() >= 3 {
                                // before or after parameter
                                let pos_val = self.evaluate_expr(&args[2])?;
                                let pos = pos_val.as_integer()? as usize;
                                // If 4th arg exists, it's "after", otherwise 3rd is "before"
                                if args.len() >= 4 {
                                    let after_val = self.evaluate_expr(&args[3])?;
                                    let after_pos = after_val.as_integer()? as usize;
                                    col_rc.borrow_mut().add_with_key_position(val, Some(&key_str), None, Some(after_pos))?;
                                } else {
                                    col_rc.borrow_mut().add_with_key_position(val, Some(&key_str), Some(pos), None)?;
                                }
                            } else {
                                col_rc.borrow_mut().add_with_key(val, &key_str)?;
                            }
                        } else {
                            col_rc.borrow_mut().add(val);
                        }
                        return Ok(Value::Nothing);
                    }
                    "remove" => {
                        let val = self.evaluate_expr(&args[0])?;
                        // If the argument is a string and it matches a key, remove by key
                        if let Value::String(key) = &val {
                            if col_rc.borrow().contains_key(key) {
                                col_rc.borrow_mut().remove_by_key(key)?;
                                return Ok(Value::Nothing);
                            }
                        }
                        // Otherwise remove by value (ArrayList semantics)
                        col_rc.borrow_mut().remove(&val);
                        return Ok(Value::Nothing);
                    }
                    "removeat" => {
                        let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                        col_rc.borrow_mut().remove_at(idx)?;
                        return Ok(Value::Nothing);
                    }
                    "clear" => {
                        col_rc.borrow_mut().clear();
                        return Ok(Value::Nothing);
                    }
                    "item" => {
                        let idx_val = self.evaluate_expr(&args[0])?;
                        match &idx_val {
                            Value::String(key) => {
                                return col_rc.borrow().item_by_key(key);
                            }
                            _ => {
                                let idx = idx_val.as_integer()? as usize;
                                return col_rc.borrow().item(idx);
                            }
                        }
                    }
                    "containskey" => {
                        let key = self.evaluate_expr(&args[0])?.as_string();
                        return Ok(Value::Boolean(col_rc.borrow().contains_key(&key)));
                    }
                    "count" => {
                        return Ok(Value::Integer(col_rc.borrow().count()));
                    }
                    "sort" => {
                        let items = &mut col_rc.borrow_mut().items;
                        items.sort_by(|a, b| {
                            let sa = a.as_string();
                            let sb = b.as_string();
                            sa.partial_cmp(&sb).unwrap_or(std::cmp::Ordering::Equal)
                        });
                        return Ok(Value::Nothing);
                    }

                    "contains" => {
                        let val = self.evaluate_expr(&args[0])?;
                        let found = col_rc.borrow().items.iter().any(|v| {
                            crate::evaluator::values_equal(v, &val)
                        });
                        return Ok(Value::Boolean(found));
                    }
                    "insert" => {
                        let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                        let val = self.evaluate_expr(&args[1])?;
                        let items = &mut col_rc.borrow_mut().items;
                        if idx <= items.len() {
                            items.insert(idx, val);
                        } else {
                            items.push(val);
                        }
                        return Ok(Value::Nothing);
                    }
                    "addrange" => {
                        let val = self.evaluate_expr(&args[0])?;
                        if let Value::Array(arr) = val {
                            for v in arr {
                                col_rc.borrow_mut().add(v);
                            }
                        } else if let Value::Collection(other) = val {
                            let items: Vec<Value> = other.borrow().items.clone();
                            for v in items {
                                col_rc.borrow_mut().add(v);
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    "find" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        for item in &items {
                            let result = self.call_lambda(predicate.clone(), &[item.clone()])?;
                            if result.as_bool()? {
                                return Ok(item.clone());
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    "findall" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        let mut result = Vec::new();
                        for item in &items {
                            let r = self.call_lambda(predicate.clone(), &[item.clone()])?;
                            if r.as_bool()? {
                                result.push(item.clone());
                            }
                        }
                        return Ok(Value::Array(result));
                    }
                    "exists" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        for item in &items {
                            let r = self.call_lambda(predicate.clone(), &[item.clone()])?;
                            if r.as_bool()? {
                                return Ok(Value::Boolean(true));
                            }
                        }
                        return Ok(Value::Boolean(false));
                    }
                    "removeall" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        let mut removed = 0i32;
                        let mut keep = Vec::new();
                        for item in &items {
                            let r = self.call_lambda(predicate.clone(), &[item.clone()])?;
                            if r.as_bool()? {
                                removed += 1;
                            } else {
                                keep.push(item.clone());
                            }
                        }
                        col_rc.borrow_mut().items = keep;
                        return Ok(Value::Integer(removed));
                    }
                    "foreach" => {
                        let action = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        for item in &items {
                            self.call_lambda(action.clone(), &[item.clone()])?;
                        }
                        return Ok(Value::Nothing);
                    }
                    "toarray" => {
                        return Ok(Value::Array(col_rc.borrow().items.clone()));
                    }
                    "lastindexof" => {
                        let val = self.evaluate_expr(&args[0])?;
                        if args.len() >= 2 {
                            let start = self.evaluate_expr(&args[1])?.as_integer()? as usize;
                            let idx = col_rc.borrow().last_index_of_from(&val, start);
                            return Ok(Value::Integer(idx));
                        }
                        let idx = col_rc.borrow().items.iter().rposition(|v| {
                            crate::evaluator::values_equal(v, &val)
                        }).map(|i| i as i32).unwrap_or(-1);
                        return Ok(Value::Integer(idx));
                    }
                    "indexof" => {
                        let val = self.evaluate_expr(&args[0])?;
                        if args.len() >= 3 {
                            let start = self.evaluate_expr(&args[1])?.as_integer()? as usize;
                            let count = self.evaluate_expr(&args[2])?.as_integer()? as usize;
                            return Ok(Value::Integer(col_rc.borrow().index_of_range(&val, start, count)));
                        } else if args.len() >= 2 {
                            let start = self.evaluate_expr(&args[1])?.as_integer()? as usize;
                            return Ok(Value::Integer(col_rc.borrow().index_of_from(&val, start)));
                        }
                        let idx = col_rc.borrow().items.iter().position(|v| {
                            crate::evaluator::values_equal(v, &val)
                        }).map(|i| i as i32).unwrap_or(-1);
                        return Ok(Value::Integer(idx));
                    }
                    "binarysearch" => {
                        let val = self.evaluate_expr(&args[0])?;
                        return Ok(Value::Integer(col_rc.borrow().binary_search(&val)));
                    }
                    "capacity" => {
                        return Ok(Value::Integer(col_rc.borrow().capacity()));
                    }
                    "trimtosize" => {
                        col_rc.borrow_mut().trim_to_size();
                        return Ok(Value::Nothing);
                    }
                    "clone" => {
                        let cloned = col_rc.borrow().clone_list();
                        return Ok(Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(cloned))));
                    }
                    "getrange" => {
                        let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                        let count = self.evaluate_expr(&args[1])?.as_integer()? as usize;
                        let range = col_rc.borrow().get_range(idx, count)?;
                        let mut al = crate::collections::ArrayList::new();
                        al.items = range;
                        return Ok(Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(al))));
                    }
                    "insertrange" => {
                        let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                        let val = self.evaluate_expr(&args[1])?;
                        let items = match val {
                            Value::Array(arr) => arr,
                            Value::Collection(c) => c.borrow().items.clone(),
                            _ => vec![val],
                        };
                        col_rc.borrow_mut().insert_range(idx, items);
                        return Ok(Value::Nothing);
                    }
                    "removerange" => {
                        let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                        let count = self.evaluate_expr(&args[1])?.as_integer()? as usize;
                        col_rc.borrow_mut().remove_range(idx, count)?;
                        return Ok(Value::Nothing);
                    }
                    "setrange" => {
                        let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                        let val = self.evaluate_expr(&args[1])?;
                        let items = match val {
                            Value::Array(arr) => arr,
                            Value::Collection(c) => c.borrow().items.clone(),
                            _ => vec![val],
                        };
                        col_rc.borrow_mut().set_range(idx, &items)?;
                        return Ok(Value::Nothing);
                    }
                    "copyto" => {
                        // CopyTo returns items as array (interpreter places into target)
                        return Ok(Value::Array(col_rc.borrow().copy_to()));
                    }
                    "findindex" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        for (i, item) in items.iter().enumerate() {
                            let r = self.call_lambda(predicate.clone(), &[item.clone()])?;
                            if r.as_bool()? {
                                return Ok(Value::Integer(i as i32));
                            }
                        }
                        return Ok(Value::Integer(-1));
                    }
                    "findlast" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        for item in items.iter().rev() {
                            let r = self.call_lambda(predicate.clone(), &[item.clone()])?;
                            if r.as_bool()? {
                                return Ok(item.clone());
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    "findlastindex" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        for (i, item) in items.iter().enumerate().rev() {
                            let r = self.call_lambda(predicate.clone(), &[item.clone()])?;
                            if r.as_bool()? {
                                return Ok(Value::Integer(i as i32));
                            }
                        }
                        return Ok(Value::Integer(-1));
                    }
                    "trueforall" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        for item in &items {
                            let r = self.call_lambda(predicate.clone(), &[item.clone()])?;
                            if !r.as_bool()? {
                                return Ok(Value::Boolean(false));
                            }
                        }
                        return Ok(Value::Boolean(true));
                    }
                    "convertall" => {
                        let converter = self.evaluate_expr(&args[0])?;
                        let items = col_rc.borrow().items.clone();
                        let mut result = Vec::new();
                        for item in &items {
                            let r = self.call_lambda(converter.clone(), &[item.clone()])?;
                            result.push(r);
                        }
                        let mut al = crate::collections::ArrayList::new();
                        al.items = result;
                        return Ok(Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(al))));
                    }
                    "reverse" => {
                        if args.len() >= 2 {
                            let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                            let count = self.evaluate_expr(&args[1])?.as_integer()? as usize;
                            col_rc.borrow_mut().reverse_range(idx, count)?;
                        } else {
                            col_rc.borrow_mut().items.reverse();
                        }
                        return Ok(Value::Nothing);
                    }
                    _ => {} // Fall through to other handlers
                 }
            }

            if let Value::Queue(q) = &obj_val {
                match method_name.as_str() {
                    "enqueue" => {
                        let val = self.evaluate_expr(&args[0])?;
                        q.borrow_mut().enqueue(val);
                        return Ok(Value::Nothing);
                    }
                    "dequeue" => {
                        return q.borrow_mut().dequeue();
                    }
                    "peek" => {
                        return q.borrow().peek();
                    }
                    "clear" => {
                        q.borrow_mut().clear();
                        return Ok(Value::Nothing);
                    }
                    "contains" => {
                        let val = self.evaluate_expr(&args[0])?;
                        return Ok(Value::Boolean(q.borrow().contains(&val)));
                    }
                    "toarray" => {
                        return Ok(Value::Array(q.borrow().to_array()));
                    }
                    "count" => {
                        return Ok(Value::Integer(q.borrow().count()));
                    }
                    _ => {}
                }
            }

            // Stack methods
            if let Value::Stack(s) = &obj_val {
                match method_name.as_str() {
                    "push" => {
                        let val = self.evaluate_expr(&args[0])?;
                        s.borrow_mut().push(val);
                        return Ok(Value::Nothing);
                    }
                    "pop" => {
                        return s.borrow_mut().pop();
                    }
                    "peek" => {
                        return s.borrow().peek();
                    }
                    "clear" => {
                        s.borrow_mut().clear();
                        return Ok(Value::Nothing);
                    }
                    "contains" => {
                        let val = self.evaluate_expr(&args[0])?;
                        return Ok(Value::Boolean(s.borrow().contains(&val)));
                    }
                    "toarray" => {
                        return Ok(Value::Array(s.borrow().to_array()));
                    }
                    "count" => {
                        return Ok(Value::Integer(s.borrow().count()));
                    }
                    _ => {}
                }
            }

            // HashSet methods
            if let Value::HashSet(h) = &obj_val {
                match method_name.as_str() {
                    "add" => {
                        let val = self.evaluate_expr(&args[0])?;
                        let was_new = h.borrow_mut().add(val);
                        return Ok(Value::Boolean(was_new));
                    }
                    "remove" => {
                        let val = self.evaluate_expr(&args[0])?;
                        let removed = h.borrow_mut().remove(&val);
                        return Ok(Value::Boolean(removed));
                    }
                    "contains" => {
                        let val = self.evaluate_expr(&args[0])?;
                        return Ok(Value::Boolean(h.borrow().contains(&val)));
                    }
                    "clear" => {
                        h.borrow_mut().clear();
                        return Ok(Value::Nothing);
                    }
                    "count" => {
                        return Ok(Value::Integer(h.borrow().count()));
                    }
                    _ => {}
                }
            }

            // Dictionary methods
            if let Value::Dictionary(d) = &obj_val {
                match method_name.as_str() {
                    "add" => {
                        let key = self.evaluate_expr(&args[0])?;
                        let val = self.evaluate_expr(&args[1])?;
                        d.borrow_mut().add(key, val)?;
                        return Ok(Value::Nothing);
                    }
                    "item" => {
                        let key = self.evaluate_expr(&args[0])?;
                        return d.borrow().item(&key);
                    }
                    "containskey" => {
                        let key = self.evaluate_expr(&args[0])?;
                        return Ok(Value::Boolean(d.borrow().contains_key(&key)));
                    }
                    "containsvalue" => {
                        let val = self.evaluate_expr(&args[0])?;
                        return Ok(Value::Boolean(d.borrow().contains_value(&val)));
                    }
                    "remove" => {
                        let key = self.evaluate_expr(&args[0])?;
                        let removed = d.borrow_mut().remove(&key);
                        return Ok(Value::Boolean(removed));
                    }
                    "clear" => {
                        d.borrow_mut().clear();
                        return Ok(Value::Nothing);
                    }
                    "count" => {
                        return Ok(Value::Integer(d.borrow().count()));
                    }
                    "keys" => {
                        let keys = d.borrow().keys();
                        return Ok(Value::Array(keys));
                    }
                    "values" => {
                        let vals = d.borrow().values();
                        return Ok(Value::Array(vals));
                    }
                    "trygetvalue" => {
                        let key = self.evaluate_expr(&args[0])?;
                        match d.borrow().item(&key) {
                            Ok(val) => {
                                // In VB.NET TryGetValue sets the ByRef param, but here
                                // we return the value. The boolean success is the wrapper.
                                // For simplicity: store value in the ByRef variable if it's a Variable expression
                                if args.len() >= 2 {
                                    if let Expression::Variable(var_name) = &args[1] {
                                        self.env.set(var_name.as_str(), val.clone()).ok();
                                    }
                                }
                                return Ok(Value::Boolean(true));
                            }
                            Err(_) => {
                                return Ok(Value::Boolean(false));
                            }
                        }
                    }
                    _ => {}
                }
            }


            // ConcurrentDictionary methods
            if let Value::ConcurrentDictionary(d) = &obj_val {
                match method_name.as_str() {
                    "tryadd" => {
                        let key = self.evaluate_expr(&args[0])?;
                        let val = self.evaluate_expr(&args[1])?;
                        return Ok(Value::Boolean(d.try_add(&key.as_string(), val)));
                    }
                    "trygetvalue" => {
                        let key = self.evaluate_expr(&args[0])?;
                        if let Some(val) = d.try_get_value(&key.as_string()) {
                             if args.len() >= 2 {
                                if let Expression::Variable(var_name) = &args[1] {
                                    self.env.set(var_name.as_str(), val.clone()).ok();
                                }
                            }
                            return Ok(Value::Boolean(true));
                        } else {
                            return Ok(Value::Boolean(false));
                        }
                    }
                    "tryremove" => {
                        let key = self.evaluate_expr(&args[0])?;
                        if let Some(val) = d.try_remove(&key.as_string()) {
                             if args.len() >= 2 {
                                if let Expression::Variable(var_name) = &args[1] {
                                    self.env.set(var_name.as_str(), val.clone()).ok();
                                }
                            }
                            return Ok(Value::Boolean(true));
                        } else {
                            return Ok(Value::Boolean(false));
                        }
                    }
                    "tryupdate" => {
                        let _key = self.evaluate_expr(&args[0])?;
                        let _new_val = self.evaluate_expr(&args[1])?;
                        let _comparison = self.evaluate_expr(&args[2])?;
                        // Comparison value also needs to be passed? ConcurrentDictionary::try_update doesn't exist in my impl?? 
                        // Wait, I didn't implement try_update in concurrent_collections.rs!
                        // I implemented add_or_update, try_add, try_get_value, try_remove, get_or_add.
                        // TryUpdate is standard?
                        // If I didn't implement it, I should remove it or implement it.
                        // I'll remove it for now or implement it if useful.
                        // Let's implement it in concurrent_collections.rs later if needed.
                        // For now, I'll comment it out or implement a dummy.
                        return Err(RuntimeError::Custom("TryUpdate not implemented".to_string()));
                    }
                    "getoradd" => {
                        let key = self.evaluate_expr(&args[0])?;
                        let val = self.evaluate_expr(&args[1])?;
                        return Ok(d.get_or_add(&key.as_string(), val));
                    }
                    "addorupdate" => {
                        let key = self.evaluate_expr(&args[0])?;
                        let add_val = self.evaluate_expr(&args[1])?;
                        let update_val = self.evaluate_expr(&args[2])?;
                        return Ok(d.add_or_update(&key.as_string(), add_val, update_val));
                    }
                    "clear" => {
                        d.clear();
                        return Ok(Value::Nothing);
                    }
                    "toarray" => {
                        // Returns KeyValuePair array
                        return Ok(Value::Array(d.to_array()));
                    }
                    _ => {}
                }

            }

            // ConcurrentQueue methods
            if let Value::ConcurrentQueue(q) = &obj_val {
                match method_name.as_str() {
                   "enqueue" => {
                       let val = self.evaluate_expr(&args[0])?;
                       q.enqueue(val);
                       return Ok(Value::Nothing);
                   }
                   "trydequeue" => {
                       if let Some(val) = q.try_dequeue() {
                            if args.len() >= 1 {
                                if let Expression::Variable(var_name) = &args[0] {
                                    self.env.set(var_name.as_str(), val.clone()).ok();
                                }
                            }
                           return Ok(Value::Boolean(true));
                       } else {
                           return Ok(Value::Boolean(false));
                       }
                   }
                   "trypeek" => {
                       if let Some(val) = q.try_peek() {
                            if args.len() >= 1 {
                                if let Expression::Variable(var_name) = &args[0] {
                                    self.env.set(var_name.as_str(), val.clone()).ok();
                                }
                            }
                           return Ok(Value::Boolean(true));
                       } else {
                           return Ok(Value::Boolean(false));
                       }
                   }
                   "toarray" => {
                       return Ok(Value::Array(q.to_array()));
                   }
                   _ => {}
                }
            }

            // ConcurrentStack methods
            if let Value::ConcurrentStack(s) = &obj_val {
                match method_name.as_str() {
                   "push" => {
                       let val = self.evaluate_expr(&args[0])?;
                       s.push(val);
                       return Ok(Value::Nothing);
                   }
                   "trypop" => {
                       if let Some(val) = s.try_pop() {
                            if args.len() >= 1 {
                                if let Expression::Variable(var_name) = &args[0] {
                                    self.env.set(var_name.as_str(), val.clone()).ok();
                                }
                            }
                           return Ok(Value::Boolean(true));
                       } else {
                           return Ok(Value::Boolean(false));
                       }
                   }
                   "trypeek" => {
                       if let Some(val) = s.try_peek() {
                            if args.len() >= 1 {
                                if let Expression::Variable(var_name) = &args[0] {
                                    self.env.set(var_name.as_str(), val.clone()).ok();
                                }
                            }
                           return Ok(Value::Boolean(true));
                       } else {
                           return Ok(Value::Boolean(false));
                       }
                   }
                   "toarray" => {
                       return Ok(Value::Array(s.to_array()));
                   }
                   _ => {}
                }
            }
        }



        // If no arguments, first try to access as a property (e.g., txt1.Text)
        if args.is_empty() {
            if let Expression::Variable(obj_name) = obj {
                // Try direct name first (e.g., btn0.Caption)
                let property_key = format!("{}.{}", obj_name.as_str(), method.as_str());
                if let Ok(val) = self.env.get(&property_key) {
                    return Ok(val);
                }
                // If the variable resolves to a string proxy, use that as the object name
                if let Ok(Value::String(proxy_name)) = self.env.get(obj_name.as_str()) {
                    let proxy_key = format!("{}.{}", proxy_name, method.as_str());
                    if let Ok(val) = self.env.get(&proxy_key) {
                        return Ok(val);
                    }
                }
            }
            
            // Allow resolving class fields/properties if parsed as MethodCall
             if let Ok(obj_val) = self.evaluate_expr(obj) {
                if let Value::Object(obj_ref) = &obj_val {
                    let class_name_str;
                    {
                        let obj_data = obj_ref.borrow();
                        class_name_str = obj_data.class_name.clone();
                        if let Some(val) = obj_data.fields.get(&method_name) {
                            return Ok(val.clone());
                        }
                    } // Drop borrow
                    
                    if let Some(prop) = self.find_property(&class_name_str, &method_name) {
                         if let Some(body) = &prop.getter {
                              let func = FunctionDecl {
                                  visibility: prop.visibility,
                                  name: prop.name.clone(),
                                  parameters: prop.parameters.clone(),
                                  return_type: prop.return_type.clone(),
                                  body: body.clone(),
                                  is_async: false,
                                  is_extension: false,
                                  is_overridable: false,
                                  is_overrides: false,
                                  is_must_override: false,
                                  is_shared: false,
                                  is_not_overridable: false,
                              };
                              return self.call_user_function(&func, &[], Some(obj_ref.clone()));
                         }
                    }
                }
             }
        }

        // Evaluate arguments
        let arg_values: Result<Vec<_>, _> = args.iter().map(|e| self.evaluate_expr(e)).collect();
        let arg_values = arg_values?;

        // Get the object name (for forms, it's usually just the identifier)
        let object_name = match obj {
            Expression::Variable(id) => id.as_str().to_string(),
            Expression::MemberAccess(_, id) => id.as_str().to_string(),
            _ => String::new(), // MethodCall or other expression (e.g. LINQ chaining)
        };

        // Check if this is a qualified procedure call (e.g., Form2.QuickTest)
        // Try looking up module.method as a Sub or Function
        let qualified_name = format!("{}.{}", object_name.to_lowercase(), method_name);
        if let Some(sub) = self.subs.get(&qualified_name).cloned() {
            return self.call_user_sub(&sub, &arg_values, None);
        }
        if let Some(func) = self.functions.get(&qualified_name).cloned() {
            return self.call_user_function(&func, &arg_values, None);
        }

        // Check if this is a class instance method call
        // We evaluate the object expression to get the instance
        if let Ok(obj_val) = self.evaluate_expr(obj) {
            // String proxy: the object is a control name string (WinForms pattern)
            // Property access like btn.Caption resolves to env key "btn0.Caption"
            if let Value::String(obj_name) = &obj_val {
                let key = format!("{}.{}", obj_name, method.as_str());
                if let Ok(val) = self.env.get(&key) {
                    return Ok(val);
                }
                // Not a known control property — fall through to other dispatch
                // (don't return empty string here, extension methods may handle it)
            }
            if let Value::Object(obj_ref) = obj_val {
                let class_name_str = obj_ref.borrow().class_name.clone();
                let class_name_lower = class_name_str.to_lowercase();

                if class_name_lower == "system.io.file" {
                    return self.dispatch_file_method(&method_name, &arg_values);
                } else if class_name_lower == "system.io.directory" {
                    return self.dispatch_directory_method(&method_name, &arg_values);
                } else if class_name_lower == "system.io.path" {
                    return self.dispatch_path_method(&method_name, &arg_values);
                } else if class_name_lower == "system.console" {
                    return self.dispatch_console_method(&method_name, &arg_values);
                } else if class_name_lower == "application" || class_name_lower == "my.application" {
                    return self.dispatch_application_method(&method_name, &arg_values);
                } else if class_name_lower == "system.drawing.color" || class_name_lower == "color" {
                    if method_name.eq_ignore_ascii_case("fromargb") {
                         return crate::builtins::drawing_fns::color_from_argb_fn(&arg_values);
                    }
                    // Standard colors
                    let argb = match method_name.to_lowercase().as_str() {
                        "red" => Some(0xFFFF0000),
                        "green" => Some(0xFF008000),
                        "blue" => Some(0xFF0000FF),
                        "black" => Some(0xFF000000),
                        "white" => Some(0xFFFFFFFF),
                        "transparent" => Some(0x00FFFFFF),
                        "yellow" => Some(0xFFFFFF00),
                        "gray" => Some(0xFF808080),
                        "lightgray" => Some(0xFFD3D3D3),
                        "darkgray" => Some(0xFFA9A9A9),
                        _ => None,
                    };
                    if let Some(val) = argb {
                        return Ok(crate::builtins::drawing_fns::create_color_obj(&method_name, val));
                    }
                } else if class_name_lower == "system.math" {
                    return self.dispatch_math_method(&method_name, &arg_values);
                } else if class_name_lower == "utf8encoding" {
                    if method_name == "getbytes" {
                        let s = arg_values.first().map(|v| v.as_string()).unwrap_or_default();
                        let bytes = s.into_bytes();
                        let value_bytes: Vec<Value> = bytes.into_iter().map(Value::Byte).collect();
                        return Ok(Value::Array(value_bytes));
                    }
                } else if class_name_lower == "bitconverter" {
                    if method_name == "tostring" {
                        let input = &arg_values[0];
                        let bytes: Vec<u8> = match input {
                            Value::Array(arr) => arr.iter().map(|v| match v {
                                Value::Byte(b) => *b,
                                Value::Integer(i) => *i as u8,
                                _ => 0u8,
                            }).collect(),
                            _ => vec![],
                        };
                        let hex: Vec<String> = bytes.iter().map(|b| format!("{:02X}", b)).collect();
                        return Ok(Value::String(hex.join("-")));
                    }
                } else if class_name_lower == "task" {
                    match method_name.as_str() {
                        "wait" => {
                            let handle_id = obj_ref.borrow().fields.get("__handle").map(|v| v.as_string()).unwrap_or_default();
                            if !handle_id.is_empty() {
                                let shared_obj_opt = get_registry().lock().unwrap().shared_objects.get(&handle_id).cloned();
                                if let Some(shared_obj) = shared_obj_opt {
                                    loop {
                                        {
                                            let lock = shared_obj.lock().unwrap();
                                            if let Some(crate::value::SharedValue::Boolean(true)) = lock.fields.get("iscompleted") {
                                                break;
                                            }
                                        }
                                        std::thread::sleep(std::time::Duration::from_millis(10));
                                    }
                                }
                            }
                            return Ok(Value::Nothing);
                        }
                        _ => {}
                    }
                } else if class_name_lower == "md5" {
                    if method_name == "computehash" {
                        return crate::builtins::cryptography_fns::md5_hash_fn(&arg_values);
                    }
                } else if class_name_lower == "sha256" {
                    if method_name == "computehash" {
                        return crate::builtins::cryptography_fns::sha256_hash_fn(&arg_values);
                    }
                } else if class_name_lower == "thread" {
                    match method_name.as_str() {
                        "start" => {
                            let task = obj_ref.borrow().fields.get("__task").cloned();
                            if let Some(lambda) = task {
                                let handle_id = generate_runtime_id();
                                obj_ref.borrow_mut().fields.insert("__handle".to_string(), Value::String(handle_id.clone()));
                                
                                let shared_lambda = lambda.to_shared();
                                
                                // Create shared object for thread state
                                let shared_thread_obj = std::sync::Arc::new(std::sync::Mutex::new(crate::value::SharedObjectData {
                                    class_name: "Thread".to_string(),
                                    fields: {
                                        let mut f = HashMap::new();
                                        f.insert("isalive".to_string(), crate::value::SharedValue::Boolean(true));
                                        f
                                    },
                                }));
                                
                                let thread_clone = shared_thread_obj.clone();
                                
                                // Clone static state for the thread
                                let functions = self.functions.clone();
                                let subs = self.subs.clone();
                                let classes = self.classes.clone();
                                let namespace_map = self.namespace_map.clone();
                                
                                let join_handle = std::thread::spawn(move || {
                                    let mut bg_interpreter = Interpreter::new_background(functions, subs, classes, namespace_map);
                                    let lambda_val = shared_lambda.to_value();
                                    let _ = bg_interpreter.call_lambda(lambda_val, &[]);
                                    
                                    // Update isalive
                                    let mut lock = thread_clone.lock().unwrap();
                                    lock.fields.insert("isalive".to_string(), crate::value::SharedValue::Boolean(false));
                                });
                                
                                {
                                    let mut reg = get_registry().lock().unwrap();
                                    reg.threads.insert(handle_id.clone(), join_handle);
                                    reg.shared_objects.insert(handle_id, shared_thread_obj);
                                }
                                
                                obj_ref.borrow_mut().fields.insert("isalive".to_string(), Value::Boolean(true));
                            }
                            return Ok(Value::Nothing);
                        }
                        "join" => {
                            let handle_id = obj_ref.borrow().fields.get("__handle").map(|v| v.as_string()).unwrap_or_default();
                            if !handle_id.is_empty() {
                                let handle = get_registry().lock().unwrap().threads.remove(&handle_id);
                                if let Some(h) = handle {
                                    let _ = h.join();
                                }
                            }
                            obj_ref.borrow_mut().fields.insert("isalive".to_string(), Value::Boolean(false));
                            return Ok(Value::Nothing);
                        }
                        _ => {}
                    }
                }
                
                // Use helper to find method in hierarchy
                if let Some(method) = self.find_method(&class_name_str, &method_name) {
                     match method {
                         vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                             self.call_user_sub(&s, &arg_values, Some(obj_ref.clone()))?;
                             return Ok(Value::Nothing);
                         }
                         vybe_parser::ast::decl::MethodDecl::Function(f) => {
                             return self.call_user_function(&f, &arg_values, Some(obj_ref.clone()));
                         }
                     }
                }
            }
        }

        // Try dispatching as a builtin qualified function call (e.g., Console.WriteLine)
        let qualified_call_name = format!("{}.{}", object_name.to_lowercase(), method_name);
        match qualified_call_name.as_str() {
            "debug.print" => {
                let msg = arg_values.iter().map(|v| v.as_string()).collect::<Vec<_>>().join(" ");
                self.send_debug_output(format!("{}\n", msg));
                return Ok(Value::Nothing);
            }
            // ---- MessageBox.Show ----
            "messagebox.show" => {
                let msg = if !arg_values.is_empty() { arg_values[0].as_string() } else { String::new() };
                let title = if arg_values.len() >= 2 { arg_values[1].as_string() } else { "vybe Basic".to_string() };
                let buttons = if arg_values.len() >= 3 {
                    match &arg_values[2] { Value::Integer(i) => *i, _ => 0 }
                } else { 0 };
                let result = crate::builtins::info_fns::show_native_msgbox(&msg, &title, buttons);
                // Return as DialogResult enum value
                return Ok(Value::Integer(result));
            }
            "console.writeline" | "console.write" | "console.readline" | "console.read"
            | "console.clear" | "console.resetcolor" | "console.beep"
            | "console.setcursorposition" => {
                let method_part = qualified_call_name.strip_prefix("console.").unwrap();
                return self.dispatch_console_method(method_part, &arg_values);
            }

            // ---- Convert class ----
            "convert.todatetime" => {
                return crate::builtins::cdate_fn(&arg_values);
            }
            "convert.toint32" | "convert.toint16" | "convert.toint64" => {
                if arg_values.len() >= 2 {
                    // Convert.ToInt32(value, base) — e.g. Convert.ToInt32("FF", 16)
                    let val_str = arg_values[0].as_string();
                    let base = arg_values[1].as_integer().unwrap_or(10);
                    match i64::from_str_radix(val_str.trim().trim_start_matches("0x").trim_start_matches("0X").trim_start_matches("&H").trim_start_matches("&h"), base as u32) {
                        Ok(n) => return Ok(Value::Integer(n as i32)),
                        Err(_) => return Err(RuntimeError::Custom(format!("Convert.ToInt32: cannot convert '{}' with base {}", val_str, base))),
                    }
                }
                return crate::builtins::cint_fn(&arg_values);
            }
            "convert.todouble" | "convert.tosingle" | "convert.todecimal" => {
                return crate::builtins::cdbl_fn(&arg_values);
            }
            "convert.tostring" => {
                if arg_values.len() >= 2 {
                    // Convert.ToString(value, base) — e.g. Convert.ToString(255, 16)
                    let val = arg_values[0].as_integer().unwrap_or(0) as i64;
                    let base = arg_values[1].as_integer().unwrap_or(10);
                    let result = match base {
                        2 => format!("{:b}", val),
                        8 => format!("{:o}", val),
                        16 => format!("{:x}", val),
                        _ => format!("{}", val),
                    };
                    return Ok(Value::String(result));
                }
                return crate::builtins::cstr_fn(&arg_values);
            }
            "convert.toboolean" => {
                return crate::builtins::cbool_fn(&arg_values);
            }
            "convert.tochar" => {
                return crate::builtins::cchar_fn(&arg_values);
            }
            "convert.tobyte" => {
                return crate::builtins::cbyte_fn(&arg_values);
            }

            // ---- Environment class ----
            "environment.getcommandlineargs" => {
                let args_array: Vec<Value> = self.command_line_args.iter()
                    .map(|a| Value::String(a.clone()))
                    .collect();
                return Ok(Value::Array(args_array));
            }
            "environment.currentdirectory" | "environment.getcurrentdirectory" => {
                let cwd = std::env::current_dir().unwrap_or_default().to_string_lossy().to_string();
                return Ok(Value::String(cwd));
            }
            "environment.machinename" => {
                let name = std::process::Command::new("hostname").output()
                    .map(|o| String::from_utf8_lossy(&o.stdout).trim().to_string())
                    .unwrap_or_else(|_| "localhost".to_string());
                return Ok(Value::String(name));
            }
            "environment.username" => {
                let name = std::env::var("USER")
                    .or_else(|_| std::env::var("USERNAME"))
                    .unwrap_or_else(|_| "unknown".to_string());
                return Ok(Value::String(name));
            }
            "environment.osversion" => {
                return Ok(Value::String(format!("{} {}", std::env::consts::OS, std::env::consts::ARCH)));
            }
            "environment.processorcount" => {
                // Use available_parallelism or fallback to 1
                let count = std::thread::available_parallelism().map(|n| n.get() as i32).unwrap_or(1);
                return Ok(Value::Integer(count));
            }
            "environment.getenvironmentvariable" => {
                let key = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                match std::env::var(&key) {
                    Ok(val) => return Ok(Value::String(val)),
                    Err(_) => return Ok(Value::Nothing),
                }
            }
            "environment.setenvironmentvariable" => {
                let key = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let val = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                // SAFETY: We are single-threaded in the VB interpreter context
                unsafe { std::env::set_var(&key, &val); }
                return Ok(Value::Nothing);
            }
            "environment.getfolderpath" => {
                // Common SpecialFolder values: Desktop=0, MyDocuments=5, AppData=26, LocalAppData=28, Temp
                let folder_id = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0);
                let path = match folder_id {
                    0 => dirs_fallback("DESKTOP", "Desktop"),       // Desktop
                    5 | 16 => dirs_fallback("HOME", "Documents"),   // MyDocuments / Personal
                    26 => std::env::var("APPDATA").unwrap_or_else(|_| {     // ApplicationData
                        let home = std::env::var("HOME").unwrap_or_default();
                        format!("{}/.config", home)
                    }),
                    28 => std::env::var("LOCALAPPDATA").unwrap_or_else(|_| { // LocalApplicationData
                        let home = std::env::var("HOME").unwrap_or_default();
                        format!("{}/.local/share", home)
                    }),
                    _ => std::env::var("HOME").unwrap_or_else(|_| "/tmp".to_string()),
                };
                return Ok(Value::String(path));
            }
            "environment.tickcount" | "environment.tickcount64" => {
                let ms = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)
                    .unwrap_or_default().as_millis() as i64;
                return Ok(Value::Long(ms));
            }
            "environment.exit" => {
                let code = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0);
                std::process::exit(code);
            }
            "environment.newline" => {
                if cfg!(target_os = "windows") {
                    return Ok(Value::String("\r\n".to_string()));
                } else {
                    return Ok(Value::String("\n".to_string()));
                }
            }
            "environment.is64bitoperatingsystem" | "environment.is64bitprocess" => {
                return Ok(Value::Boolean(cfg!(target_pointer_width = "64")));
            }
            "environment.version" => {
                return Ok(Value::String("4.0.0".to_string())); // simulate .NET 4
            }

            // ---- Guid class ----
            "guid.newguid" | "system.guid.newguid" => {
                // Generate a UUID v4 using timestamp + counter
                let now = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default();
                let nanos = now.as_nanos();
                let a = ((nanos >> 96) as u32) ^ (nanos as u32);
                let b = ((nanos >> 64) as u16) ^ ((nanos >> 16) as u16);
                let c = (((nanos >> 48) as u16) & 0x0FFF) | 0x4000; // version 4
                let d = (((nanos >> 32) as u8) & 0x3F) | 0x80; // variant
                let e = (nanos >> 24) as u8;
                let rest = nanos as u64;
                let guid_str = format!("{:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:012x}",
                    a, b, c, d, e, rest & 0xFFFFFFFFFFFF);
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("Guid".to_string()));
                fields.insert("__value".to_string(), Value::String(guid_str.clone()));
                let obj = crate::value::ObjectData { class_name: "Guid".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }
            "guid.parse" | "system.guid.parse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("Guid".to_string()));
                fields.insert("__value".to_string(), Value::String(s));
                let obj = crate::value::ObjectData { class_name: "Guid".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }
            "guid.empty" | "system.guid.empty" => {
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("Guid".to_string()));
                fields.insert("__value".to_string(), Value::String("00000000-0000-0000-0000-000000000000".to_string()));
                let obj = crate::value::ObjectData { class_name: "Guid".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }

            // ---- Convert Base64 ----
            "convert.tobase64string" => {
                let input = &arg_values[0];
                let bytes: Vec<u8> = match input {
                    Value::Array(arr) => arr.iter().map(|v| match v {
                        Value::Byte(b) => *b,
                        Value::Integer(i) => *i as u8,
                        _ => 0u8,
                    }).collect(),
                    Value::String(s) => s.as_bytes().to_vec(),
                    _ => vec![],
                };
                // Manual base64 encode
                let encoded = base64_encode(&bytes);
                return Ok(Value::String(encoded));
            }
            "convert.frombase64string" => {
                let s = arg_values[0].as_string();
                match base64_decode(&s) {
                    Ok(bytes) => {
                        // Return as string (common usage) — VB.NET returns Byte() but
                        // most callers immediately convert to string
                        let decoded = String::from_utf8(bytes.clone()).unwrap_or_else(|_| {
                            // If not valid UTF-8, return byte array
                            return bytes.iter().map(|b| *b as char).collect();
                        });
                        return Ok(Value::String(decoded));
                    }
                    Err(e) => return Err(RuntimeError::Custom(format!("Convert.FromBase64String: {}", e))),
                }
            }

            // ---- Stopwatch.StartNew (factory) ----
            "stopwatch.startnew" | "system.diagnostics.stopwatch.startnew" => {
                let now_ms = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_millis() as i64;
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("Stopwatch".to_string()));
                fields.insert("isrunning".to_string(), Value::Boolean(true));
                fields.insert("elapsedmilliseconds".to_string(), Value::Long(0));
                fields.insert("__start_ms".to_string(), Value::Long(now_ms));
                fields.insert("__accumulated_ms".to_string(), Value::Long(0));
                let obj = crate::value::ObjectData { class_name: "Stopwatch".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }

            // ---- Thread.Sleep ----
            "thread.sleep" | "system.threading.thread.sleep" => {
                let ms = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0);
                std::thread::sleep(std::time::Duration::from_millis(ms.max(0) as u64));
                return Ok(Value::Nothing);
            }

            // ===== TASK STATIC METHODS =====
            "task.run" | "system.threading.tasks.task.run" => {
                let lambda = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                let handle_id = generate_runtime_id();
                
                // Create shared task object
                let mut fields = HashMap::new();
                fields.insert("__type".to_string(), crate::value::SharedValue::String("Task".to_string()));
                fields.insert("iscompleted".to_string(), crate::value::SharedValue::Boolean(false));
                fields.insert("result".to_string(), crate::value::SharedValue::Nothing);
                fields.insert("__handle".to_string(), crate::value::SharedValue::String(handle_id.clone()));
                fields.insert("status".to_string(), crate::value::SharedValue::String("WaitingToRun".to_string()));
                
                let shared_task_obj = std::sync::Arc::new(std::sync::Mutex::new(crate::value::SharedObjectData {
                    class_name: "Task".to_string(),
                    fields,
                }));
                
                let shared_lambda = lambda.to_shared();
                let functions = self.functions.clone();
                let subs = self.subs.clone();
                let classes = self.classes.clone();
                let namespace_map = self.namespace_map.clone();
                
                let task_clone = shared_task_obj.clone();
                
                let join_handle = std::thread::spawn(move || {
                    let mut bg_interpreter = Interpreter::new_background(functions, subs, classes, namespace_map);
                    let result = match bg_interpreter.call_lambda(shared_lambda.to_value(), &[]) {
                        Ok(v) => v.to_shared(),
                        Err(_) => crate::value::SharedValue::Nothing,
                    };
                    
                    let mut lock = task_clone.lock().unwrap();
                    lock.fields.insert("iscompleted".to_string(), crate::value::SharedValue::Boolean(true));
                    lock.fields.insert("result".to_string(), result);
                    lock.fields.insert("status".to_string(), crate::value::SharedValue::String("RanToCompletion".to_string()));
                });
                
                {
                    let mut reg = get_registry().lock().unwrap();
                    reg.threads.insert(handle_id.clone(), join_handle);
                    reg.shared_objects.insert(handle_id, shared_task_obj.clone());
                }
                
                return Ok(crate::value::SharedValue::Object(shared_task_obj).to_value());
            }
            "task.delay" | "system.threading.tasks.task.delay" => {
                let ms = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0) as u64;
                let handle_id = generate_runtime_id();
                
                let shared_task_obj = std::sync::Arc::new(std::sync::Mutex::new(crate::value::SharedObjectData {
                    class_name: "Task".to_string(),
                    fields: {
                        let mut f = HashMap::new();
                        f.insert("__handle".to_string(), crate::value::SharedValue::String(handle_id.clone()));
                        f.insert("iscompleted".to_string(), crate::value::SharedValue::Boolean(false));
                        f.insert("status".to_string(), crate::value::SharedValue::String("Running".to_string()));
                        f.insert("result".to_string(), crate::value::SharedValue::Nothing);
                        f
                    },
                }));
                
                let task_clone = shared_task_obj.clone();
                let join_handle = std::thread::spawn(move || {
                    std::thread::sleep(std::time::Duration::from_millis(ms));
                    let mut lock = task_clone.lock().unwrap();
                    lock.fields.insert("iscompleted".to_string(), crate::value::SharedValue::Boolean(true));
                    lock.fields.insert("status".to_string(), crate::value::SharedValue::String("RanToCompletion".to_string()));
                });
                
                {
                    let mut reg = get_registry().lock().unwrap();
                    reg.threads.insert(handle_id.clone(), join_handle);
                    reg.shared_objects.insert(handle_id, shared_task_obj.clone());
                }
                
                return Ok(crate::value::SharedValue::Object(shared_task_obj).to_value());
            }
            "task.fromresult" | "system.threading.tasks.task.fromresult" => {
                let val = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("Task".to_string()));
                fields.insert("result".to_string(), val);
                fields.insert("iscompleted".to_string(), Value::Boolean(true));
                fields.insert("isfaulted".to_string(), Value::Boolean(false));
                fields.insert("iscanceled".to_string(), Value::Boolean(false));
                fields.insert("status".to_string(), Value::String("RanToCompletion".to_string()));
                let obj = crate::value::ObjectData { class_name: "Task".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }
            "task.whenall" | "system.threading.tasks.task.whenall" => {
                // All tasks are already completed (synchronous), just return completed task
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("Task".to_string()));
                fields.insert("result".to_string(), Value::Nothing);
                fields.insert("iscompleted".to_string(), Value::Boolean(true));
                fields.insert("isfaulted".to_string(), Value::Boolean(false));
                fields.insert("iscanceled".to_string(), Value::Boolean(false));
                fields.insert("status".to_string(), Value::String("RanToCompletion".to_string()));
                let obj = crate::value::ObjectData { class_name: "Task".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }
            "task.whenany" | "system.threading.tasks.task.whenany" => {
                let first = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("Task".to_string()));
                fields.insert("result".to_string(), first);
                fields.insert("iscompleted".to_string(), Value::Boolean(true));
                fields.insert("isfaulted".to_string(), Value::Boolean(false));
                fields.insert("iscanceled".to_string(), Value::Boolean(false));
                fields.insert("status".to_string(), Value::String("RanToCompletion".to_string()));
                let obj = crate::value::ObjectData { class_name: "Task".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }
            "task.completedtask" | "system.threading.tasks.task.completedtask" => {
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("Task".to_string()));
                fields.insert("result".to_string(), Value::Nothing);
                fields.insert("iscompleted".to_string(), Value::Boolean(true));
                fields.insert("isfaulted".to_string(), Value::Boolean(false));
                fields.insert("iscanceled".to_string(), Value::Boolean(false));
                fields.insert("status".to_string(), Value::String("RanToCompletion".to_string()));
                let obj = crate::value::ObjectData { class_name: "Task".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }

            // ===== THREADPOOL =====
            "threadpool.queueuserworkitem" | "system.threading.threadpool.queueuserworkitem" => {
                if let Some(lambda) = arg_values.get(0).cloned() {
                    let state = arg_values.get(1).cloned().unwrap_or(Value::Nothing);
                    
                    let shared_lambda = lambda.to_shared();
                    let shared_state = state.to_shared();
                    
                    let functions = self.functions.clone();
                    let subs = self.subs.clone();
                    let classes = self.classes.clone();
                    let namespace_map = self.namespace_map.clone();
                    
                    std::thread::spawn(move || {
                        let mut bg_interpreter = Interpreter::new_background(functions, subs, classes, namespace_map);
                        let _ = bg_interpreter.call_lambda(shared_lambda.to_value(), &[shared_state.to_value()]);
                    });
                    return Ok(Value::Boolean(true));
                }
                return Ok(Value::Boolean(false));
            }
            "threadpool.setminthreads" | "system.threading.threadpool.setminthreads" => {
                return Ok(Value::Boolean(true)); // No-op, always succeed
            }
            "threadpool.setmaxthreads" | "system.threading.threadpool.setmaxthreads" => {
                return Ok(Value::Boolean(true));
            }

            // ===== INTERLOCKED (SYNCHRONOUS — single-threaded interpreter) =====
            "interlocked.increment" | "system.threading.interlocked.increment" => {
                // Interlocked.Increment(ByRef variable) — increment and return new value
                // In vybe, args[0] is a Variable expression that was already evaluated
                // We need the variable name to update it. Use a fallback: get the value, increment, return.
                if let Some(Expression::Variable(var_name)) = args.get(0) {
                    let current = self.env.get(var_name.as_str()).unwrap_or(Value::Integer(0));
                    let new_val = match current {
                        Value::Integer(i) => Value::Integer(i + 1),
                        Value::Long(l) => Value::Long(l + 1),
                        _ => Value::Integer(current.as_integer().unwrap_or(0) + 1),
                    };
                    let _ = self.env.set(var_name.as_str(), new_val.clone());
                    return Ok(new_val);
                }
                // Fallback: just return arg + 1
                let val = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0);
                return Ok(Value::Integer(val + 1));
            }
            "interlocked.decrement" | "system.threading.interlocked.decrement" => {
                if let Some(Expression::Variable(var_name)) = args.get(0) {
                    let current = self.env.get(var_name.as_str()).unwrap_or(Value::Integer(0));
                    let new_val = match current {
                        Value::Integer(i) => Value::Integer(i - 1),
                        Value::Long(l) => Value::Long(l - 1),
                        _ => Value::Integer(current.as_integer().unwrap_or(0) - 1),
                    };
                    let _ = self.env.set(var_name.as_str(), new_val.clone());
                    return Ok(new_val);
                }
                let val = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0);
                return Ok(Value::Integer(val - 1));
            }
            "interlocked.exchange" | "system.threading.interlocked.exchange" => {
                // Interlocked.Exchange(ByRef location, value) — sets location = value, returns old
                if let Some(Expression::Variable(var_name)) = args.get(0) {
                    let old_val = self.env.get(var_name.as_str()).unwrap_or(Value::Nothing);
                    let new_val = arg_values.get(1).cloned().unwrap_or(Value::Nothing);
                    let _ = self.env.set(var_name.as_str(), new_val);
                    return Ok(old_val);
                }
                return Ok(arg_values.get(0).cloned().unwrap_or(Value::Nothing));
            }
            "interlocked.compareexchange" | "system.threading.interlocked.compareexchange" => {
                // Interlocked.CompareExchange(ByRef location, value, comparand)
                // If location == comparand, sets location = value. Returns old value of location.
                if let Some(Expression::Variable(var_name)) = args.get(0) {
                    let old_val = self.env.get(var_name.as_str()).unwrap_or(Value::Nothing);
                    let new_val = arg_values.get(1).cloned().unwrap_or(Value::Nothing);
                    let comparand = arg_values.get(2).cloned().unwrap_or(Value::Nothing);
                    if old_val == comparand {
                        let _ = self.env.set(var_name.as_str(), new_val);
                    }
                    return Ok(old_val);
                }
                return Ok(arg_values.get(0).cloned().unwrap_or(Value::Nothing));
            }

            // ===== ACTIVATOR.CREATEINSTANCE =====
            "activator.createinstance" | "system.activator.createinstance" => {
                // Activator.CreateInstance(typeName) — create an instance of a class by name
                let type_name = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let type_lower = type_name.to_lowercase();
                // Check if it's a known class
                if let Some(class) = self.classes.get(&type_lower).cloned() {
                    // Create an instance using the class definition
                    let mut fields = std::collections::HashMap::new();
                    let class_name_str = class.name.as_str().to_string();
                    fields.insert("__type".to_string(), Value::String(class_name_str.clone()));
                    // Initialize fields from class definition
                    for field in &class.fields {
                        let field_name = field.name.as_str().to_lowercase();
                        let default_val = if let Some(init) = &field.initializer {
                            self.evaluate_expr(init).unwrap_or(Value::Nothing)
                        } else {
                            Value::Nothing
                        };
                        fields.insert(field_name, default_val);
                    }
                    let obj = crate::value::ObjectData { class_name: class_name_str.clone(), fields };
                    let obj_ref = std::rc::Rc::new(std::cell::RefCell::new(obj));

                    // Call Sub New if it exists
                    if let Some(new_method) = self.find_method(&class_name_str, "new") {
                         match new_method {
                             vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                                 // Activator.CreateInstance(type) calls parameterless New()
                                 // Or matching arguments if provided (but here we only implement 1-arg version)
                                 // Assuming parameterless for now.
                                 let _ = self.call_user_sub(&s, &[], Some(obj_ref.clone()));
                             }
                             _ => {}
                         }
                    }
                    
                    return Ok(Value::Object(obj_ref));
                }
                // Check built-in types
                match type_lower.as_str() {
                    "arraylist" | "system.collections.arraylist" => {
                        return Ok(Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(
                            crate::collections::ArrayList::new()
                        ))));
                    }
                    "dictionary" | "system.collections.generic.dictionary" | "hashtable" => {
                        return Ok(Value::Dictionary(std::rc::Rc::new(std::cell::RefCell::new(
                            crate::collections::VBDictionary::new()
                        ))));
                    }
                    "stringbuilder" | "system.text.stringbuilder" => {
                        let mut fields = std::collections::HashMap::new();
                        fields.insert("__type".to_string(), Value::String("StringBuilder".to_string()));
                        fields.insert("__buffer".to_string(), Value::String(String::new()));
                        let obj = crate::value::ObjectData { class_name: "StringBuilder".to_string(), fields };
                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                    }
                    _ => {
                        // Create a generic object
                        let mut fields = std::collections::HashMap::new();
                        fields.insert("__type".to_string(), Value::String(type_name.clone()));
                        let obj = crate::value::ObjectData { class_name: type_name, fields };
                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                    }
                }
            }

            // ===== CALLBYNAME =====
            "callbyname" | "microsoft.visualbasic.interaction.callbyname" => {
                // CallByName(obj, memberName, callType, args...)
                // callType: 1=Method, 2=Get, 4=Set, 8=Let
                let obj_val = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                let member = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                let call_type = arg_values.get(2).map(|v| v.as_integer().unwrap_or(1)).unwrap_or(1);
                let member_lower = member.to_lowercase();

                if let Value::Object(obj_ref) = &obj_val {
                    match call_type {
                        2 | 8 => {
                            // Get property
                            let obj_data = obj_ref.borrow();
                            if let Some(val) = obj_data.fields.get(&member_lower) {
                                return Ok(val.clone());
                            }
                            return Ok(Value::Nothing);
                        }
                        4 => {
                            // Set property
                            let set_val = arg_values.get(3).cloned().unwrap_or(Value::Nothing);
                            obj_ref.borrow_mut().fields.insert(member_lower, set_val);
                            return Ok(Value::Nothing);
                        }
                        _ => {
                            // Method call (call_type = 1)
                            let class_name = obj_ref.borrow().class_name.clone();
                            // Check for a sub or function with name "classname.member"
                            let qualified = format!("{}.{}", class_name.to_lowercase(), member_lower);
                            let method_args: Vec<Value> = arg_values[3..].to_vec();
                            if let Some(sub) = self.subs.get(&qualified).cloned() {
                                return self.call_user_sub(&sub, &method_args, Some(obj_ref.clone()));
                            }
                            if let Some(func) = self.functions.get(&qualified).cloned() {
                                return self.call_user_function(&func, &method_args, Some(obj_ref.clone()));
                            }
                            // Try direct method name
                            if let Some(sub) = self.subs.get(&member_lower).cloned() {
                                return self.call_user_sub(&sub, &method_args, Some(obj_ref.clone()));
                            }
                            if let Some(func) = self.functions.get(&member_lower).cloned() {
                                return self.call_user_function(&func, &method_args, Some(obj_ref.clone()));
                            }
                            return Err(RuntimeError::Custom(format!("CallByName: method '{}' not found on '{}'", member, class_name)));
                        }
                    }
                }
                // For non-object types, try to call as a method on the value
                if call_type == 2 || call_type == 8 {
                    // Get — try as_string on common types
                    return Ok(Value::Nothing);
                }
                return Err(RuntimeError::Custom(format!("CallByName: target is not an object")));
            }

            // ===== TUPLE.CREATE =====
            "tuple.create" | "system.tuple.create" => {
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("Tuple".to_string()));
                for (i, val) in arg_values.iter().enumerate() {
                    fields.insert(format!("item{}", i + 1), val.clone());
                }
                fields.insert("length".to_string(), Value::Integer(arg_values.len() as i32));
                let obj = crate::value::ObjectData { class_name: "Tuple".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }

            // ===== ZIPFILE STATIC METHODS =====
            "zipfile.createfromdirectory" | "system.io.compression.zipfile.createfromdirectory" => {
                let src_dir = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let dest_zip = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                // Use system zip command
                let status = std::process::Command::new("zip")
                    .args(&["-r", &dest_zip, "."])
                    .current_dir(&src_dir)
                    .status();
                match status {
                    Ok(s) if s.success() => return Ok(Value::Nothing),
                    _ => return Err(RuntimeError::Custom(format!("ZipFile.CreateFromDirectory failed"))),
                }
            }
            "zipfile.extracttodirectory" | "system.io.compression.zipfile.extracttodirectory" => {
                let src_zip = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let dest_dir = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                let _ = std::fs::create_dir_all(&dest_dir);
                let status = std::process::Command::new("unzip")
                    .args(&["-o", &src_zip, "-d", &dest_dir])
                    .status();
                match status {
                    Ok(s) if s.success() => return Ok(Value::Nothing),
                    _ => return Err(RuntimeError::Custom(format!("ZipFile.ExtractToDirectory failed"))),
                }
            }
            "zipfile.open" | "system.io.compression.zipfile.open" => {
                let path = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let mode = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("ZipArchive".to_string()));
                fields.insert("__path".to_string(), Value::String(path));
                fields.insert("__mode".to_string(), Value::Integer(mode));
                let obj = crate::value::ObjectData { class_name: "ZipArchive".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }

            // ===== BITCONVERTER STATIC METHODS =====
            "bitconverter.getbytes" | "system.bitconverter.getbytes" => {
                let val = &arg_values[0];
                let bytes = match val {
                    Value::Integer(i) => i.to_le_bytes().to_vec(),
                    Value::Long(l) => l.to_le_bytes().to_vec(),
                    Value::Double(d) => d.to_le_bytes().to_vec(),
                    Value::Single(f) => f.to_le_bytes().to_vec(),
                    Value::Boolean(b) => vec![if *b { 1u8 } else { 0u8 }],
                    _ => vec![],
                };
                return Ok(Value::Array(bytes.iter().map(|b| Value::Integer(*b as i32)).collect()));
            }
            "bitconverter.toint32" | "system.bitconverter.toint32" => {
                if let Value::Array(arr) = &arg_values[0] {
                    let start = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                    let bytes: Vec<u8> = arr[start..start+4.min(arr.len()-start)].iter()
                        .map(|v| v.as_integer().unwrap_or(0) as u8).collect();
                    if bytes.len() >= 4 {
                        return Ok(Value::Integer(i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]])));
                    }
                }
                return Ok(Value::Integer(0));
            }
            "bitconverter.toint64" | "system.bitconverter.toint64" => {
                if let Value::Array(arr) = &arg_values[0] {
                    let start = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                    let bytes: Vec<u8> = arr[start..start+8.min(arr.len()-start)].iter()
                        .map(|v| v.as_integer().unwrap_or(0) as u8).collect();
                    if bytes.len() >= 8 {
                        return Ok(Value::Long(i64::from_le_bytes([bytes[0],bytes[1],bytes[2],bytes[3],bytes[4],bytes[5],bytes[6],bytes[7]])));
                    }
                }
                return Ok(Value::Long(0));
            }
            "bitconverter.todouble" | "system.bitconverter.todouble" => {
                if let Value::Array(arr) = &arg_values[0] {
                    let start = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                    let bytes: Vec<u8> = arr[start..start+8.min(arr.len()-start)].iter()
                        .map(|v| v.as_integer().unwrap_or(0) as u8).collect();
                    if bytes.len() >= 8 {
                        return Ok(Value::Double(f64::from_le_bytes([bytes[0],bytes[1],bytes[2],bytes[3],bytes[4],bytes[5],bytes[6],bytes[7]])));
                    }
                }
                return Ok(Value::Double(0.0));
            }
            "bitconverter.tostring" | "system.bitconverter.tostring" => {
                if let Value::Array(arr) = &arg_values[0] {
                    let hex: Vec<String> = arr.iter().map(|v| format!("{:02X}", v.as_integer().unwrap_or(0) as u8)).collect();
                    return Ok(Value::String(hex.join("-")));
                }
                return Ok(Value::String(String::new()));
            }

            // ---- Process.Start ----
            "process.start" | "system.diagnostics.process.start" => {
                let mut filename = String::new();
                let mut arguments = String::new();
                
                if let Some(val) = arg_values.get(0) {
                    match val {
                        Value::Object(obj_ref) => {
                            let b = obj_ref.borrow();
                            if b.class_name.eq_ignore_ascii_case("ProcessStartInfo") {
                                filename = b.fields.get("filename").map(|v| v.as_string()).unwrap_or_default();
                                arguments = b.fields.get("arguments").map(|v| v.as_string()).unwrap_or_default();
                            } else {
                                filename = val.as_string();
                            }
                        }
                        _ => {
                            filename = val.as_string();
                            arguments = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                        }
                    }
                }

                if filename.is_empty() {
                    return Err(RuntimeError::Custom("Process.Start: missing executable path".to_string()));
                }

                let mut cmd = std::process::Command::new(&filename);
                if !arguments.is_empty() {
                    // Split args by space (simple split)
                    for a in arguments.split_whitespace() {
                        cmd.arg(a);
                    }
                }
                match cmd.spawn() {
                    Ok(child) => {
                        // Generate a unique handle ID
                        let handle_id = format!("proc_{}", child.id());
                        
                        // Register the child process in the global registry
                        {
                            let mut reg = get_registry().lock().unwrap();
                            reg.processes.insert(handle_id.clone(), child);
                        }

                        // Return a Process object with Id and __handle
                        let mut fields = std::collections::HashMap::new();
                        fields.insert("__type".to_string(), Value::String("Process".to_string()));
                        fields.insert("__handle".to_string(), Value::String(handle_id));
                        fields.insert("id".to_string(), Value::Integer(0)); // child.id() is used for handle
                        fields.insert("hasexited".to_string(), Value::Boolean(false));
                        fields.insert("exitcode".to_string(), Value::Integer(0));
                        
                        let obj = crate::value::ObjectData { class_name: "Process".to_string(), fields };
                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                    }
                    Err(e) => return Err(RuntimeError::Custom(format!("Process.Start failed: {} (file={})", e, filename))),
                }
            }

            // ---- Debug.Write / Debug.WriteLine / Debug.Assert ----
            "debug.write" | "system.diagnostics.debug.write" => {
                let msg = arg_values.iter().map(|v| v.as_string()).collect::<Vec<_>>().join(" ");
                self.send_debug_output(msg);
                return Ok(Value::Nothing);
            }
            "debug.writeline" | "system.diagnostics.debug.writeline" => {
                let msg = arg_values.iter().map(|v| v.as_string()).collect::<Vec<_>>().join(" ");
                self.send_debug_output(format!("{}\n", msg));
                return Ok(Value::Nothing);
            }
            "debug.assert" | "system.diagnostics.debug.assert" => {
                let condition = arg_values.get(0).map(|v| v.is_truthy()).unwrap_or(true);
                if !condition {
                    let msg = arg_values.get(1).map(|v| v.as_string()).unwrap_or_else(|| "Debug.Assert failed".to_string());
                    self.send_debug_output(format!("ASSERT FAILED: {}\n", msg));
                }
                return Ok(Value::Nothing);
            }

            // ---- String static methods ----
            "string.format" | "system.string.format" => {
                // String.Format("{0} is {1}", arg0, arg1)
                let fmt = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let mut result = fmt;
                for (i, arg) in arg_values.iter().skip(1).enumerate() {
                    result = result.replace(&format!("{{{}}}", i), &arg.as_string());
                }
                return Ok(Value::String(result));
            }
            "string.join" | "system.string.join" => {
                // String.Join(separator, array_or_items...)
                let sep = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                if let Some(Value::Array(arr)) = arg_values.get(1) {
                    let joined = arr.iter().map(|v| v.as_string()).collect::<Vec<_>>().join(&sep);
                    return Ok(Value::String(joined));
                } else {
                    // Join remaining args
                    let joined = arg_values.iter().skip(1).map(|v| v.as_string()).collect::<Vec<_>>().join(&sep);
                    return Ok(Value::String(joined));
                }
            }
            "string.isnullorwhitespace" | "system.string.isnullorwhitespace" => {
                let s = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                let is_empty = match &s {
                    Value::Nothing => true,
                    Value::String(st) => st.trim().is_empty(),
                    _ => false,
                };
                return Ok(Value::Boolean(is_empty));
            }
            "string.concat" | "system.string.concat" => {
                let result: String = arg_values.iter().map(|v| v.as_string()).collect();
                return Ok(Value::String(result));
            }
            "string.compare" | "system.string.compare" => {
                let a = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let b = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                let ignore_case = arg_values.get(2).map(|v| v.is_truthy()).unwrap_or(false);
                let cmp = if ignore_case {
                    a.to_lowercase().cmp(&b.to_lowercase())
                } else {
                    a.cmp(&b)
                };
                return Ok(Value::Integer(match cmp {
                    std::cmp::Ordering::Less => -1,
                    std::cmp::Ordering::Equal => 0,
                    std::cmp::Ordering::Greater => 1,
                }));
            }
            "string.empty" | "system.string.empty" => {
                return Ok(Value::String(String::new()));
            }
            "eventargs.empty" | "system.eventargs.empty" => {
                return Ok(Self::make_event_args());
            }
            "string.isnullorempty" | "system.string.isnullorempty" => {
                let s = arg_values.get(0).cloned().unwrap_or(Value::Nothing);
                let is_empty = match &s {
                    Value::Nothing => true,
                    Value::String(st) => st.is_empty(),
                    _ => false,
                };
                return Ok(Value::Boolean(is_empty));
            }

            // ---- Type.Parse static methods ----
            "integer.parse" | "int32.parse" | "system.int32.parse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                match s.trim().parse::<i32>() {
                    Ok(n) => return Ok(Value::Integer(n)),
                    Err(_) => return Err(RuntimeError::Custom(format!("Integer.Parse: '{}' is not a valid integer", s))),
                }
            }
            "integer.tryparse" | "int32.tryparse" | "system.int32.tryparse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                return Ok(Value::Boolean(s.trim().parse::<i32>().is_ok()));
            }
            "long.parse" | "int64.parse" | "system.int64.parse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                match s.trim().parse::<i64>() {
                    Ok(n) => return Ok(Value::Long(n)),
                    Err(_) => return Err(RuntimeError::Custom(format!("Long.Parse: '{}' is not a valid long", s))),
                }
            }
            "long.tryparse" | "int64.tryparse" | "system.int64.tryparse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                return Ok(Value::Boolean(s.trim().parse::<i64>().is_ok()));
            }
            "double.parse" | "system.double.parse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                match s.trim().parse::<f64>() {
                    Ok(n) => return Ok(Value::Double(n)),
                    Err(_) => return Err(RuntimeError::Custom(format!("Double.Parse: '{}' is not a valid double", s))),
                }
            }
            "double.tryparse" | "system.double.tryparse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                return Ok(Value::Boolean(s.trim().parse::<f64>().is_ok()));
            }
            "single.parse" | "system.single.parse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                match s.trim().parse::<f32>() {
                    Ok(n) => return Ok(Value::Single(n)),
                    Err(_) => return Err(RuntimeError::Custom(format!("Single.Parse: '{}' is not a valid single", s))),
                }
            }
            "decimal.parse" | "system.decimal.parse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                match s.trim().parse::<f64>() {
                    Ok(n) => return Ok(Value::Double(n)),
                    Err(_) => return Err(RuntimeError::Custom(format!("Decimal.Parse: '{}' is not a valid decimal", s))),
                }
            }
            "boolean.parse" | "system.boolean.parse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                match s.trim().to_lowercase().as_str() {
                    "true" => return Ok(Value::Boolean(true)),
                    "false" => return Ok(Value::Boolean(false)),
                    _ => return Err(RuntimeError::Custom(format!("Boolean.Parse: '{}' is not a valid Boolean", s))),
                }
            }

            // ---- Array static methods ----
            "array.sort" | "system.array.sort" => {
                if let Some(Value::Array(arr)) = arg_values.get(0) {
                    let mut sorted = arr.clone();
                    sorted.sort_by(|a, b| a.as_string().cmp(&b.as_string()));
                    return Ok(Value::Array(sorted));
                }
                return Err(RuntimeError::Custom("Array.Sort requires an array argument".to_string()));
            }
            "array.reverse" | "system.array.reverse" => {
                if let Some(Value::Array(arr)) = arg_values.get(0) {
                    let mut reversed = arr.clone();
                    reversed.reverse();
                    return Ok(Value::Array(reversed));
                }
                return Err(RuntimeError::Custom("Array.Reverse requires an array argument".to_string()));
            }
            "array.indexof" | "system.array.indexof" => {
                if let Some(Value::Array(arr)) = arg_values.get(0) {
                    let needle = arg_values.get(1).cloned().unwrap_or(Value::Nothing);
                    let needle_str = needle.as_string();
                    let idx = arr.iter().position(|v| v.as_string() == needle_str);
                    return Ok(Value::Integer(idx.map(|i| i as i32).unwrap_or(-1)));
                }
                return Err(RuntimeError::Custom("Array.IndexOf requires an array argument".to_string()));
            }
            "array.find" | "system.array.find" => {
                // Array.Find(array, predicate) — for now, just return Nothing as lambdas need special handling
                // Try to execute lambda if provided
                if let (Some(Value::Array(arr)), Some(lambda_val)) = (arg_values.get(0), arg_values.get(1)) {
                    if matches!(lambda_val, Value::Lambda { .. }) {
                        for item in arr {
                            let result = self.call_lambda(lambda_val.clone(), &[item.clone()])?;
                            if result.is_truthy() {
                                return Ok(item.clone());
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                }
                return Ok(Value::Nothing);
            }
            "array.findall" | "system.array.findall" => {
                if let (Some(Value::Array(arr)), Some(lambda_val)) = (arg_values.get(0), arg_values.get(1)) {
                    if matches!(lambda_val, Value::Lambda { .. }) {
                        let mut results = Vec::new();
                        for item in arr {
                            let result = self.call_lambda(lambda_val.clone(), &[item.clone()])?;
                            if result.is_truthy() {
                                results.push(item.clone());
                            }
                        }
                        return Ok(Value::Array(results));
                    }
                }
                return Ok(Value::Array(Vec::new()));
            }
            "array.copy" | "system.array.copy" => {
                // Array.Copy(source, dest, length) — returns new array with copied elements
                if let Some(Value::Array(src)) = arg_values.get(0) {
                    let len = arg_values.get(2).map(|v| v.as_integer().unwrap_or(src.len() as i32)).unwrap_or(src.len() as i32) as usize;
                    let copied: Vec<Value> = src.iter().take(len).cloned().collect();
                    return Ok(Value::Array(copied));
                }
                return Err(RuntimeError::Custom("Array.Copy requires a source array".to_string()));
            }
            "array.resize" | "system.array.resize" => {
                // Array.Resize(ByRef array, newSize) — returns new array
                if let Some(Value::Array(arr)) = arg_values.get(0) {
                    let new_size = arg_values.get(1).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0) as usize;
                    let mut resized = arr.clone();
                    resized.resize(new_size, Value::Nothing);
                    return Ok(Value::Array(resized));
                }
                return Err(RuntimeError::Custom("Array.Resize requires an array argument".to_string()));
            }
            "array.clear" | "system.array.clear" => {
                // Array.Clear(array, index, length) — returns array with portion cleared
                if let Some(Value::Array(arr)) = arg_values.get(0) {
                    let start = arg_values.get(1).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0) as usize;
                    let length = arg_values.get(2).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0) as usize;
                    let mut cleared = arr.clone();
                    for i in start..std::cmp::min(start + length, cleared.len()) {
                        cleared[i] = Value::Nothing;
                    }
                    return Ok(Value::Array(cleared));
                }
                return Err(RuntimeError::Custom("Array.Clear requires an array argument".to_string()));
            }
            "array.exists" | "system.array.exists" => {
                if let (Some(Value::Array(arr)), Some(lambda_val)) = (arg_values.get(0), arg_values.get(1)) {
                    if matches!(lambda_val, Value::Lambda { .. }) {
                        for item in arr {
                            let result = self.call_lambda(lambda_val.clone(), &[item.clone()])?;
                            if result.is_truthy() {
                                return Ok(Value::Boolean(true));
                            }
                        }
                    }
                }
                return Ok(Value::Boolean(false));
            }

            // ---- DateTime static methods ----
            "datetime.now" | "system.datetime.now" => {
                return Ok(Value::Date(now_ole()));
            }
            "datetime.today" | "system.datetime.today" => {
                return Ok(Value::Date(today_ole()));
            }
            "datetime.utcnow" | "system.datetime.utcnow" => {
                return Ok(Value::Date(utcnow_ole()));
            }
            "datetime.parse" | "system.datetime.parse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                match parse_date_to_ole(&s) {
                    Some(ole) => return Ok(Value::Date(ole)),
                    None => return Err(RuntimeError::Custom(format!("DateTime.Parse: cannot parse '{}'", s))),
                }
            }
            "datetime.tryparse" | "system.datetime.tryparse" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                return Ok(Value::Boolean(parse_date_to_ole(&s).is_some()));
            }
            "datetime.daysinmonth" | "system.datetime.daysinmonth" => {
                let year = arg_values.get(0).map(|v| v.as_integer().unwrap_or(2024)).unwrap_or(2024);
                let month = arg_values.get(1).map(|v| v.as_integer().unwrap_or(1)).unwrap_or(1);
                // Calculate days in month
                let days = if month == 12 {
                    31
                } else {
                    let d1 = chrono::NaiveDate::from_ymd_opt(year, month as u32, 1);
                    let d2 = chrono::NaiveDate::from_ymd_opt(if month == 12 { year + 1 } else { year }, if month == 12 { 1 } else { (month + 1) as u32 }, 1);
                    match (d1, d2) {
                        (Some(a), Some(b)) => (b - a).num_days() as i32,
                        _ => 30,
                    }
                };
                return Ok(Value::Integer(days));
            }
            "datetime.isleapyear" | "system.datetime.isleapyear" => {
                let year = arg_values.get(0).map(|v| v.as_integer().unwrap_or(2024)).unwrap_or(2024);
                let leap = (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
                return Ok(Value::Boolean(leap));
            }

            // ---- TimeSpan additional factory methods ----
            "timespan.frommilliseconds" => {
                let ms = arg_values[0].as_double()?;
                let total_seconds = ms / 1000.0;
                let obj_data = crate::value::ObjectData {
                    class_name: "TimeSpan".to_string(),
                    fields: {
                        let mut f = std::collections::HashMap::new();
                        f.insert("days".to_string(), Value::Integer(0));
                        f.insert("hours".to_string(), Value::Integer(0));
                        f.insert("minutes".to_string(), Value::Integer(0));
                        f.insert("seconds".to_string(), Value::Integer(0));
                        f.insert("milliseconds".to_string(), Value::Integer(ms as i32));
                        f.insert("totaldays".to_string(), Value::Double(total_seconds / 86400.0));
                        f.insert("totalhours".to_string(), Value::Double(total_seconds / 3600.0));
                        f.insert("totalminutes".to_string(), Value::Double(total_seconds / 60.0));
                        f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                        f.insert("totalmilliseconds".to_string(), Value::Double(ms));
                        f
                    },
                };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
            }
            "timespan.fromticks" => {
                let ticks = arg_values[0].as_double()?;
                let ms = ticks / 10000.0;
                let total_seconds = ms / 1000.0;
                let obj_data = crate::value::ObjectData {
                    class_name: "TimeSpan".to_string(),
                    fields: {
                        let mut f = std::collections::HashMap::new();
                        f.insert("days".to_string(), Value::Integer(0));
                        f.insert("hours".to_string(), Value::Integer(0));
                        f.insert("minutes".to_string(), Value::Integer(0));
                        f.insert("seconds".to_string(), Value::Integer(0));
                        f.insert("milliseconds".to_string(), Value::Integer(ms as i32));
                        f.insert("ticks".to_string(), Value::Long(ticks as i64));
                        f.insert("totaldays".to_string(), Value::Double(total_seconds / 86400.0));
                        f.insert("totalhours".to_string(), Value::Double(total_seconds / 3600.0));
                        f.insert("totalminutes".to_string(), Value::Double(total_seconds / 60.0));
                        f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                        f.insert("totalmilliseconds".to_string(), Value::Double(ms));
                        f
                    },
                };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
            }
            "timespan.parse" | "system.timespan.parse" => {
                // Parse "hh:mm:ss" or "d.hh:mm:ss" format
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let parts: Vec<&str> = s.split(':').collect();
                let (hours, minutes, seconds) = match parts.len() {
                    3 => {
                        let h: f64 = parts[0].parse().unwrap_or(0.0);
                        let m: f64 = parts[1].parse().unwrap_or(0.0);
                        let sec: f64 = parts[2].parse().unwrap_or(0.0);
                        (h, m, sec)
                    }
                    2 => {
                        let m: f64 = parts[0].parse().unwrap_or(0.0);
                        let sec: f64 = parts[1].parse().unwrap_or(0.0);
                        (0.0, m, sec)
                    }
                    _ => (0.0, 0.0, 0.0),
                };
                let total_seconds = hours * 3600.0 + minutes * 60.0 + seconds;
                let obj_data = crate::value::ObjectData {
                    class_name: "TimeSpan".to_string(),
                    fields: {
                        let mut f = std::collections::HashMap::new();
                        f.insert("days".to_string(), Value::Integer((total_seconds / 86400.0) as i32));
                        f.insert("hours".to_string(), Value::Integer(hours as i32));
                        f.insert("minutes".to_string(), Value::Integer(minutes as i32));
                        f.insert("seconds".to_string(), Value::Integer(seconds as i32));
                        f.insert("milliseconds".to_string(), Value::Integer(0));
                        f.insert("totaldays".to_string(), Value::Double(total_seconds / 86400.0));
                        f.insert("totalhours".to_string(), Value::Double(total_seconds / 3600.0));
                        f.insert("totalminutes".to_string(), Value::Double(total_seconds / 60.0));
                        f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                        f.insert("totalmilliseconds".to_string(), Value::Double(total_seconds * 1000.0));
                        f
                    },
                };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
            }
            "timespan.zero" | "system.timespan.zero" => {
                let obj_data = crate::value::ObjectData {
                    class_name: "TimeSpan".to_string(),
                    fields: {
                        let mut f = std::collections::HashMap::new();
                        f.insert("days".to_string(), Value::Integer(0));
                        f.insert("hours".to_string(), Value::Integer(0));
                        f.insert("minutes".to_string(), Value::Integer(0));
                        f.insert("seconds".to_string(), Value::Integer(0));
                        f.insert("milliseconds".to_string(), Value::Integer(0));
                        f.insert("totaldays".to_string(), Value::Double(0.0));
                        f.insert("totalhours".to_string(), Value::Double(0.0));
                        f.insert("totalminutes".to_string(), Value::Double(0.0));
                        f.insert("totalseconds".to_string(), Value::Double(0.0));
                        f.insert("totalmilliseconds".to_string(), Value::Double(0.0));
                        f
                    },
                };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
            }

            // ---- Uri class ----
            "uri.iswell formeduristring" | "uri.iswellformeduristring" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let is_valid = s.starts_with("http://") || s.starts_with("https://") || s.starts_with("ftp://") || s.starts_with("file://");
                return Ok(Value::Boolean(is_valid));
            }
            "uri.escapedatastring" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let encoded: String = s.chars().map(|c| {
                    if c.is_ascii_alphanumeric() || c == '-' || c == '_' || c == '.' || c == '~' {
                        c.to_string()
                    } else {
                        format!("%{:02X}", c as u32)
                    }
                }).collect();
                return Ok(Value::String(encoded));
            }
            "uri.unescapedatastring" => {
                let s = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                // Simple percent-decode
                let mut result = String::new();
                let mut chars = s.chars();
                while let Some(c) = chars.next() {
                    if c == '%' {
                        let hex: String = chars.by_ref().take(2).collect();
                        if let Ok(byte) = u8::from_str_radix(&hex, 16) {
                            result.push(byte as char);
                        }
                    } else {
                        result.push(c);
                    }
                }
                return Ok(Value::String(result));
            }

            // ---- Console.ReadKey ----
            "console.readkey" => {
                // In GUI/CLI context, return a ConsoleKeyInfo-like object
                let mut fields = std::collections::HashMap::new();
                fields.insert("keychar".to_string(), Value::String(" ".to_string()));
                fields.insert("key".to_string(), Value::Integer(32)); // Spacebar
                let obj = crate::value::ObjectData { class_name: "ConsoleKeyInfo".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }

            // ---- WebRequest.Create ----
            "webrequest.create" | "system.net.webrequest.create"
            | "httpwebrequest.create" | "system.net.httpwebrequest.create" => {
                let url = arg_values[0].as_string();
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("HttpWebRequest".to_string()));
                fields.insert("url".to_string(), Value::String(url));
                fields.insert("method".to_string(), Value::String("GET".to_string()));
                fields.insert("contenttype".to_string(), Value::String("application/x-www-form-urlencoded".to_string()));
                fields.insert("useragent".to_string(), Value::String("vybe/1.0".to_string()));
                fields.insert("timeout".to_string(), Value::Integer(100000));
                fields.insert("headers".to_string(), Value::Dictionary(
                    std::rc::Rc::new(std::cell::RefCell::new(crate::collections::VBDictionary::new()))
                ));
                let obj = crate::value::ObjectData { class_name: "HttpWebRequest".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }

            // ---- ServicePointManager (no-op, just absorb assignments) ----
            "servicepointmanager.securityprotocol" => {
                // No-op: TLS is handled by curl
                return Ok(Value::Integer(0));
            }

            // ---- Dns class ----
            "dns.gethostentry" | "system.net.dns.gethostentry" => {
                let host = arg_values[0].as_string();
                let output = std::process::Command::new("host")
                    .arg(&host)
                    .output();
                let mut addresses = Vec::new();
                let mut hostname = host.clone();
                if let Ok(out) = output {
                    let text = String::from_utf8_lossy(&out.stdout).to_string();
                    for line in text.lines() {
                        if line.contains("has address") {
                            if let Some(addr) = line.split_whitespace().last() {
                                addresses.push(Value::String(addr.to_string()));
                            }
                        }
                        if line.contains("has IPv6 address") {
                            if let Some(addr) = line.split_whitespace().last() {
                                addresses.push(Value::String(addr.to_string()));
                            }
                        }
                        // Extract canonical hostname
                        if line.contains("is an alias for") {
                            if let Some(alias) = line.split("is an alias for ").nth(1) {
                                hostname = alias.trim_end_matches('.').to_string();
                            }
                        }
                    }
                }
                // Fallback: try getent / dig if host not available
                if addresses.is_empty() {
                    if let Ok(out) = std::process::Command::new("getent").args(&["hosts", &host]).output() {
                        let text = String::from_utf8_lossy(&out.stdout).to_string();
                        for line in text.lines() {
                            if let Some(addr) = line.split_whitespace().next() {
                                addresses.push(Value::String(addr.to_string()));
                            }
                        }
                    }
                }
                let mut fields = std::collections::HashMap::new();
                fields.insert("hostname".to_string(), Value::String(hostname));
                fields.insert("addresslist".to_string(), Value::Array(addresses));
                let obj = crate::value::ObjectData { class_name: "IPHostEntry".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }
            "dns.gethostaddresses" | "system.net.dns.gethostaddresses" => {
                let host = arg_values[0].as_string();
                let output = std::process::Command::new("host")
                    .arg(&host)
                    .output();
                let mut addresses = Vec::new();
                if let Ok(out) = output {
                    let text = String::from_utf8_lossy(&out.stdout).to_string();
                    for line in text.lines() {
                        if line.contains("has address") || line.contains("has IPv6 address") {
                            if let Some(addr) = line.split_whitespace().last() {
                                addresses.push(Value::String(addr.to_string()));
                            }
                        }
                    }
                }
                if addresses.is_empty() {
                    if let Ok(out) = std::process::Command::new("getent").args(&["hosts", &host]).output() {
                        let text = String::from_utf8_lossy(&out.stdout).to_string();
                        for line in text.lines() {
                            if let Some(addr) = line.split_whitespace().next() {
                                addresses.push(Value::String(addr.to_string()));
                            }
                        }
                    }
                }
                return Ok(Value::Array(addresses));
            }
            "dns.gethostname" | "system.net.dns.gethostname" => {
                let output = std::process::Command::new("hostname").output();
                let name = if let Ok(out) = output {
                    String::from_utf8_lossy(&out.stdout).trim().to_string()
                } else {
                    "localhost".to_string()
                };
                return Ok(Value::String(name));
            }

            // ---- IPAddress class ----
            "ipaddress.parse" | "system.net.ipaddress.parse" => {
                let addr_str = arg_values[0].as_string();
                // Validate: try to parse as IPv4 or IPv6
                let is_v4 = addr_str.parse::<std::net::Ipv4Addr>().is_ok();
                let is_v6 = addr_str.parse::<std::net::Ipv6Addr>().is_ok();
                if !is_v4 && !is_v6 {
                    return Err(RuntimeError::Custom(format!("IPAddress.Parse: invalid address '{}'", addr_str)));
                }
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("IPAddress".to_string()));
                fields.insert("address".to_string(), Value::String(addr_str.clone()));
                fields.insert("addressfamily".to_string(), Value::String(
                    if is_v4 { "InterNetwork".to_string() } else { "InterNetworkV6".to_string() }
                ));
                let obj = crate::value::ObjectData { class_name: "IPAddress".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }
            "ipaddress.tryparse" | "system.net.ipaddress.tryparse" => {
                let addr_str = arg_values[0].as_string();
                let is_v4 = addr_str.parse::<std::net::Ipv4Addr>().is_ok();
                let is_v6 = addr_str.parse::<std::net::Ipv6Addr>().is_ok();
                if is_v4 || is_v6 {
                    // Store the parsed address in the ByRef second arg if possible
                    // For now, just return True (the parsed value can be re-obtained via Parse)
                    return Ok(Value::Boolean(true));
                }
                return Ok(Value::Boolean(false));
            }
            "ipaddress.loopback" | "system.net.ipaddress.loopback" => {
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("IPAddress".to_string()));
                fields.insert("address".to_string(), Value::String("127.0.0.1".to_string()));
                fields.insert("addressfamily".to_string(), Value::String("InterNetwork".to_string()));
                let obj = crate::value::ObjectData { class_name: "IPAddress".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }
            "ipaddress.any" | "system.net.ipaddress.any" => {
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("IPAddress".to_string()));
                fields.insert("address".to_string(), Value::String("0.0.0.0".to_string()));
                fields.insert("addressfamily".to_string(), Value::String("InterNetwork".to_string()));
                let obj = crate::value::ObjectData { class_name: "IPAddress".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }
            "ipaddress.ipv6loopback" | "system.net.ipaddress.ipv6loopback" => {
                let mut fields = std::collections::HashMap::new();
                fields.insert("__type".to_string(), Value::String("IPAddress".to_string()));
                fields.insert("address".to_string(), Value::String("::1".to_string()));
                fields.insert("addressfamily".to_string(), Value::String("InterNetworkV6".to_string()));
                let obj = crate::value::ObjectData { class_name: "IPAddress".to_string(), fields };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
            }

            // ---- TimeSpan factory methods ----
            "timespan.fromdays" => {
                let days = arg_values[0].as_double()?;
                let total_seconds = days * 86400.0;
                let obj_data = crate::value::ObjectData {
                    class_name: "TimeSpan".to_string(),
                    fields: {
                        let mut f = std::collections::HashMap::new();
                        f.insert("days".to_string(), Value::Integer(days as i32));
                        f.insert("hours".to_string(), Value::Integer(0));
                        f.insert("minutes".to_string(), Value::Integer(0));
                        f.insert("seconds".to_string(), Value::Integer(0));
                        f.insert("milliseconds".to_string(), Value::Integer(0));
                        f.insert("totaldays".to_string(), Value::Double(days));
                        f.insert("totalhours".to_string(), Value::Double(days * 24.0));
                        f.insert("totalminutes".to_string(), Value::Double(days * 1440.0));
                        f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                        f.insert("totalmilliseconds".to_string(), Value::Double(total_seconds * 1000.0));
                        f
                    },
                };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
            }
            "timespan.fromhours" => {
                let hours = arg_values[0].as_double()?;
                let total_seconds = hours * 3600.0;
                let obj_data = crate::value::ObjectData {
                    class_name: "TimeSpan".to_string(),
                    fields: {
                        let mut f = std::collections::HashMap::new();
                        f.insert("days".to_string(), Value::Integer(0));
                        f.insert("hours".to_string(), Value::Integer(hours as i32));
                        f.insert("minutes".to_string(), Value::Integer(0));
                        f.insert("seconds".to_string(), Value::Integer(0));
                        f.insert("totaldays".to_string(), Value::Double(hours / 24.0));
                        f.insert("totalhours".to_string(), Value::Double(hours));
                        f.insert("totalminutes".to_string(), Value::Double(hours * 60.0));
                        f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                        f.insert("totalmilliseconds".to_string(), Value::Double(total_seconds * 1000.0));
                        f
                    },
                };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
            }
            "timespan.fromminutes" => {
                let mins = arg_values[0].as_double()?;
                let total_seconds = mins * 60.0;
                let obj_data = crate::value::ObjectData {
                    class_name: "TimeSpan".to_string(),
                    fields: {
                        let mut f = std::collections::HashMap::new();
                        f.insert("days".to_string(), Value::Integer(0));
                        f.insert("hours".to_string(), Value::Integer(0));
                        f.insert("minutes".to_string(), Value::Integer(mins as i32));
                        f.insert("seconds".to_string(), Value::Integer(0));
                        f.insert("totaldays".to_string(), Value::Double(mins / 1440.0));
                        f.insert("totalhours".to_string(), Value::Double(mins / 60.0));
                        f.insert("totalminutes".to_string(), Value::Double(mins));
                        f.insert("totalseconds".to_string(), Value::Double(total_seconds));
                        f.insert("totalmilliseconds".to_string(), Value::Double(total_seconds * 1000.0));
                        f
                    },
                };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
            }
            "timespan.fromseconds" => {
                let secs = arg_values[0].as_double()?;
                let obj_data = crate::value::ObjectData {
                    class_name: "TimeSpan".to_string(),
                    fields: {
                        let mut f = std::collections::HashMap::new();
                        f.insert("days".to_string(), Value::Integer(0));
                        f.insert("hours".to_string(), Value::Integer(0));
                        f.insert("minutes".to_string(), Value::Integer(0));
                        f.insert("seconds".to_string(), Value::Integer(secs as i32));
                        f.insert("totaldays".to_string(), Value::Double(secs / 86400.0));
                        f.insert("totalhours".to_string(), Value::Double(secs / 3600.0));
                        f.insert("totalminutes".to_string(), Value::Double(secs / 60.0));
                        f.insert("totalseconds".to_string(), Value::Double(secs));
                        f.insert("totalmilliseconds".to_string(), Value::Double(secs * 1000.0));
                        f
                    },
                };
                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj_data))));
            }

            _ => {
                // Prefix-based dispatch for static class calls (Path.*, Math.*, File.*, Console.*)
                if qualified_call_name.starts_with("path.") || qualified_call_name.starts_with("system.io.path.") {
                    return self.dispatch_path_method(&method_name, &arg_values);
                } else if qualified_call_name.starts_with("math.") || qualified_call_name.starts_with("system.math.") {
                    return self.dispatch_math_method(&method_name, &arg_values);
                } else if qualified_call_name.starts_with("file.") || qualified_call_name.starts_with("system.io.file.") {
                    return self.dispatch_file_method(&method_name, &arg_values);
                } else if qualified_call_name.starts_with("directory.") || qualified_call_name.starts_with("system.io.directory.") {
                    return self.dispatch_directory_method(&method_name, &arg_values);
                } else if qualified_call_name.starts_with("messagebox.") || qualified_call_name.starts_with("system.windows.forms.messagebox.") {
                    if method_name.eq_ignore_ascii_case("show") {
                        return crate::builtins::msgbox::msgbox(&arg_values);
                    }
                } else if qualified_call_name.starts_with("color.") || qualified_call_name.starts_with("system.drawing.color.") {
                    if method_name.eq_ignore_ascii_case("fromargb") {
                         return crate::builtins::drawing_fns::color_from_argb_fn(&arg_values);
                    }
                }
            }
        }

        // ── LINQ-style extension methods ────────────────────────────────────
        // These work on any iterable value (Array, Collection, Dictionary,
        // Queue, Stack, HashSet, String) via to_iterable().
        // Re-use the already-evaluated obj_val from above to avoid double evaluation
        // (which breaks chaining like .Where().Select()).
        if let Ok(ref val) = eval_result {
            if let Ok(items) = val.to_iterable() {
                match method_name.as_str() {
                    // .Select(Function(x) expr)
                    "select" => {
                        let selector = self.evaluate_expr(&args[0])?;
                        let mut result = Vec::new();
                        for item in &items {
                            result.push(self.call_lambda(selector.clone(), &[item.clone()])?);
                        }
                        return Ok(Value::Array(result));
                    }
                    // .Where(Function(x) bool_expr)
                    "where" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let mut result = Vec::new();
                        for item in &items {
                            let r = self.call_lambda(predicate.clone(), &[item.clone()])?;
                            if r.is_truthy() {
                                result.push(item.clone());
                            }
                        }
                        return Ok(Value::Array(result));
                    }
                    // .First() / .First(Function(x) bool)
                    "first" => {
                        if args.is_empty() {
                            return items.first().cloned().ok_or_else(|| {
                                RuntimeError::Custom("Sequence contains no elements".to_string())
                            });
                        }
                        let predicate = self.evaluate_expr(&args[0])?;
                        for item in &items {
                            if self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                return Ok(item.clone());
                            }
                        }
                        return Err(RuntimeError::Custom("Sequence contains no matching element".to_string()));
                    }
                    // .FirstOrDefault() / .FirstOrDefault(Function(x) bool)
                    "firstordefault" => {
                        if args.is_empty() {
                            return Ok(items.first().cloned().unwrap_or(Value::Nothing));
                        }
                        let predicate = self.evaluate_expr(&args[0])?;
                        for item in &items {
                            if self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                return Ok(item.clone());
                            }
                        }
                        return Ok(Value::Nothing);
                    }
                    // .Last() / .Last(Function(x) bool)
                    "last" => {
                        if args.is_empty() {
                            return items.last().cloned().ok_or_else(|| {
                                RuntimeError::Custom("Sequence contains no elements".to_string())
                            });
                        }
                        let predicate = self.evaluate_expr(&args[0])?;
                        let mut found = None;
                        for item in &items {
                            if self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                found = Some(item.clone());
                            }
                        }
                        return found.ok_or_else(|| RuntimeError::Custom("Sequence contains no matching element".to_string()));
                    }
                    // .LastOrDefault()
                    "lastordefault" => {
                        if args.is_empty() {
                            return Ok(items.last().cloned().unwrap_or(Value::Nothing));
                        }
                        let predicate = self.evaluate_expr(&args[0])?;
                        let mut found = Value::Nothing;
                        for item in &items {
                            if self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                found = item.clone();
                            }
                        }
                        return Ok(found);
                    }
                    // .Single() / .Single(Function(x) bool)
                    "single" => {
                        let filtered: Vec<_> = if args.is_empty() {
                            items.clone()
                        } else {
                            let predicate = self.evaluate_expr(&args[0])?;
                            let mut r = Vec::new();
                            for item in &items {
                                if self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                    r.push(item.clone());
                                }
                            }
                            r
                        };
                        match filtered.len() {
                            0 => return Err(RuntimeError::Custom("Sequence contains no matching element".to_string())),
                            1 => return Ok(filtered[0].clone()),
                            _ => return Err(RuntimeError::Custom("Sequence contains more than one matching element".to_string())),
                        }
                    }
                    "singleordefault" => {
                        let filtered: Vec<_> = if args.is_empty() {
                            items.clone()
                        } else {
                            let predicate = self.evaluate_expr(&args[0])?;
                            let mut r = Vec::new();
                            for item in &items {
                                if self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                    r.push(item.clone());
                                }
                            }
                            r
                        };
                        match filtered.len() {
                            0 => return Ok(Value::Nothing),
                            1 => return Ok(filtered[0].clone()),
                            _ => return Err(RuntimeError::Custom("Sequence contains more than one matching element".to_string())),
                        }
                    }
                    // .Count() / .Count(Function(x) bool)
                    "count" if !matches!(val, Value::Object(_)) => {
                        if args.is_empty() {
                            return Ok(Value::Integer(items.len() as i32));
                        }
                        let predicate = self.evaluate_expr(&args[0])?;
                        let mut count = 0i32;
                        for item in &items {
                            if self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                count += 1;
                            }
                        }
                        return Ok(Value::Integer(count));
                    }
                    // .Any() / .Any(Function(x) bool)
                    "any" => {
                        if args.is_empty() {
                            return Ok(Value::Boolean(!items.is_empty()));
                        }
                        let predicate = self.evaluate_expr(&args[0])?;
                        for item in &items {
                            if self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                return Ok(Value::Boolean(true));
                            }
                        }
                        return Ok(Value::Boolean(false));
                    }
                    // .All(Function(x) bool)
                    "all" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        for item in &items {
                            if !self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                return Ok(Value::Boolean(false));
                            }
                        }
                        return Ok(Value::Boolean(true));
                    }
                    // .OrderBy(Function(x) key) / .OrderByDescending(Function(x) key)
                    "orderby" | "orderbydescending" => {
                        let selector = self.evaluate_expr(&args[0])?;
                        let mut keyed: Vec<(Value, Value)> = Vec::new();
                        for item in &items {
                            let key = self.call_lambda(selector.clone(), &[item.clone()])?;
                            keyed.push((item.clone(), key));
                        }
                        keyed.sort_by(|a, b| compare_values_ordering(&a.1, &b.1));
                        if method_name == "orderbydescending" {
                            keyed.reverse();
                        }
                        return Ok(Value::Array(keyed.into_iter().map(|(v, _)| v).collect()));
                    }
                    // .Skip(n)
                    "skip" => {
                        let n = self.evaluate_expr(&args[0])?.as_integer()?.max(0) as usize;
                        return Ok(Value::Array(items.into_iter().skip(n).collect()));
                    }
                    // .Take(n)
                    "take" => {
                        let n = self.evaluate_expr(&args[0])?.as_integer()?.max(0) as usize;
                        return Ok(Value::Array(items.into_iter().take(n).collect()));
                    }
                    // .Distinct()
                    "distinct" => {
                        let mut seen = Vec::new();
                        let mut result = Vec::new();
                        for item in &items {
                            let key = item.as_string();
                            if !seen.contains(&key) {
                                seen.push(key);
                                result.push(item.clone());
                            }
                        }
                        return Ok(Value::Array(result));
                    }
                    // .Sum() / .Sum(Function(x) num)
                    "sum" => {
                        let mut total = 0.0f64;
                        if args.is_empty() {
                            for item in &items { total += item.as_double().unwrap_or(0.0); }
                        } else {
                            let selector = self.evaluate_expr(&args[0])?;
                            for item in &items {
                                let v = self.call_lambda(selector.clone(), &[item.clone()])?;
                                total += v.as_double().unwrap_or(0.0);
                            }
                        }
                        return Ok(Value::Double(total));
                    }
                    // .Min() / .Max()
                    "min" => {
                        if items.is_empty() { return Err(RuntimeError::Custom("Sequence contains no elements".to_string())); }
                        if args.is_empty() {
                            let mut min_val = items[0].as_double().unwrap_or(f64::MAX);
                            for item in &items[1..] { min_val = min_val.min(item.as_double().unwrap_or(f64::MAX)); }
                            return Ok(Value::Double(min_val));
                        }
                        let selector = self.evaluate_expr(&args[0])?;
                        let mut min_val = f64::MAX;
                        for item in &items {
                            let v = self.call_lambda(selector.clone(), &[item.clone()])?.as_double().unwrap_or(f64::MAX);
                            min_val = min_val.min(v);
                        }
                        return Ok(Value::Double(min_val));
                    }
                    "max" => {
                        if items.is_empty() { return Err(RuntimeError::Custom("Sequence contains no elements".to_string())); }
                        if args.is_empty() {
                            let mut max_val = items[0].as_double().unwrap_or(f64::MIN);
                            for item in &items[1..] { max_val = max_val.max(item.as_double().unwrap_or(f64::MIN)); }
                            return Ok(Value::Double(max_val));
                        }
                        let selector = self.evaluate_expr(&args[0])?;
                        let mut max_val = f64::MIN;
                        for item in &items {
                            let v = self.call_lambda(selector.clone(), &[item.clone()])?.as_double().unwrap_or(f64::MIN);
                            max_val = max_val.max(v);
                        }
                        return Ok(Value::Double(max_val));
                    }
                    // .Average() / .Average(Function(x) num)
                    "average" => {
                        if items.is_empty() { return Err(RuntimeError::Custom("Sequence contains no elements".to_string())); }
                        let mut total = 0.0f64;
                        if args.is_empty() {
                            for item in &items { total += item.as_double().unwrap_or(0.0); }
                        } else {
                            let selector = self.evaluate_expr(&args[0])?;
                            for item in &items {
                                let v = self.call_lambda(selector.clone(), &[item.clone()])?;
                                total += v.as_double().unwrap_or(0.0);
                            }
                        }
                        return Ok(Value::Double(total / items.len() as f64));
                    }
                    // .Aggregate(seed, Function(acc, x) expr)
                    "aggregate" => {
                        if args.len() >= 2 {
                            let mut acc = self.evaluate_expr(&args[0])?;
                            let func = self.evaluate_expr(&args[1])?;
                            for item in &items {
                                acc = self.call_lambda(func.clone(), &[acc, item.clone()])?;
                            }
                            return Ok(acc);
                        } else if args.len() == 1 && !items.is_empty() {
                            let func = self.evaluate_expr(&args[0])?;
                            let mut acc = items[0].clone();
                            for item in &items[1..] {
                                acc = self.call_lambda(func.clone(), &[acc, item.clone()])?;
                            }
                            return Ok(acc);
                        }
                        return Err(RuntimeError::Custom("Aggregate requires at least one argument".to_string()));
                    }
                    // .Reverse()
                    "reverse" => {
                        let mut rev = items;
                        rev.reverse();
                        return Ok(Value::Array(rev));
                    }
                    // .Contains(value) — generic LINQ Contains
                    "contains" if !matches!(val, Value::Object(_)) => {
                        let target = self.evaluate_expr(&args[0])?;
                        let target_str = target.as_string();
                        let found = items.iter().any(|v| v.as_string() == target_str);
                        return Ok(Value::Boolean(found));
                    }
                    // .ToList() — returns Collection/ArrayList
                    "tolist" => {
                        let al = crate::collections::ArrayList { items, keys: std::collections::HashMap::new() };
                        return Ok(Value::Collection(std::rc::Rc::new(std::cell::RefCell::new(al))));
                    }
                    // .ToArray()
                    "toarray" if !matches!(val, Value::Object(_)) => {
                        return Ok(Value::Array(items));
                    }
                    // .ToDictionary(Function(x) key, Function(x) val)
                    "todictionary" => {
                        let key_sel = self.evaluate_expr(&args[0])?;
                        let val_sel = if args.len() > 1 { self.evaluate_expr(&args[1])? } else { Value::Nothing };
                        let mut dict = crate::collections::VBDictionary::new();
                        for item in &items {
                            let k = self.call_lambda(key_sel.clone(), &[item.clone()])?;
                            let v = if matches!(val_sel, Value::Nothing) {
                                item.clone()
                            } else {
                                self.call_lambda(val_sel.clone(), &[item.clone()])?
                            };
                            let _ = dict.add(k, v);
                        }
                        return Ok(Value::Dictionary(std::rc::Rc::new(std::cell::RefCell::new(dict))));
                    }
                    // .SelectMany(Function(x) array)
                    "selectmany" => {
                        let selector = self.evaluate_expr(&args[0])?;
                        let mut result = Vec::new();
                        for item in &items {
                            let sub = self.call_lambda(selector.clone(), &[item.clone()])?;
                            if let Ok(sub_items) = sub.to_iterable() {
                                result.extend(sub_items);
                            } else {
                                result.push(sub);
                            }
                        }
                        return Ok(Value::Array(result));
                    }
                    // .Zip(other, Function(a, b) expr)
                    "zip" => {
                        let other_val = self.evaluate_expr(&args[0])?;
                        let other_items = other_val.to_iterable()?;
                        let combiner = if args.len() > 1 { Some(self.evaluate_expr(&args[1])?) } else { None };
                        let mut result = Vec::new();
                        for (a, b) in items.iter().zip(other_items.iter()) {
                            if let Some(ref func) = combiner {
                                result.push(self.call_lambda(func.clone(), &[a.clone(), b.clone()])?);
                            } else {
                                // Return as tuple-like object
                                let mut fields = std::collections::HashMap::new();
                                fields.insert("item1".to_string(), a.clone());
                                fields.insert("item2".to_string(), b.clone());
                                fields.insert("__type".to_string(), Value::String("Tuple".to_string()));
                                result.push(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(
                                    crate::value::ObjectData { class_name: "Tuple".to_string(), fields }
                                ))));
                            }
                        }
                        return Ok(Value::Array(result));
                    }
                    // .GroupBy(Function(x) key) → array of group objects {Key, Items}
                    "groupby" => {
                        let key_sel = self.evaluate_expr(&args[0])?;
                        let mut groups: Vec<(String, Vec<Value>)> = Vec::new();
                        for item in &items {
                            let key = self.call_lambda(key_sel.clone(), &[item.clone()])?.as_string();
                            if let Some(g) = groups.iter_mut().find(|(k, _)| k == &key) {
                                g.1.push(item.clone());
                            } else {
                                groups.push((key, vec![item.clone()]));
                            }
                        }
                        let mut result = Vec::new();
                        for (key, group_items) in groups {
                            let mut fields = std::collections::HashMap::new();
                            fields.insert("key".to_string(), Value::String(key));
                            fields.insert("items".to_string(), Value::Array(group_items.clone()));
                            fields.insert("count".to_string(), Value::Integer(group_items.len() as i32));
                            fields.insert("__type".to_string(), Value::String("Grouping".to_string()));
                            result.push(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(
                                crate::value::ObjectData { class_name: "Grouping".to_string(), fields }
                            ))));
                        }
                        return Ok(Value::Array(result));
                    }
                    // .Union(other) — set union preserving order
                    "union" => {
                        let other = self.evaluate_expr(&args[0])?;
                        let other_items = other.to_iterable().unwrap_or_default();
                        let mut seen = Vec::new();
                        let mut result = Vec::new();
                        for item in items.iter().chain(other_items.iter()) {
                            let key = item.as_string();
                            if !seen.contains(&key) {
                                seen.push(key);
                                result.push(item.clone());
                            }
                        }
                        return Ok(Value::Array(result));
                    }
                    // .Intersect(other) — set intersection
                    "intersect" => {
                        let other = self.evaluate_expr(&args[0])?;
                        let other_items = other.to_iterable().unwrap_or_default();
                        let other_keys: Vec<String> = other_items.iter().map(|v| v.as_string()).collect();
                        let mut seen = Vec::new();
                        let mut result = Vec::new();
                        for item in &items {
                            let key = item.as_string();
                            if other_keys.contains(&key) && !seen.contains(&key) {
                                seen.push(key);
                                result.push(item.clone());
                            }
                        }
                        return Ok(Value::Array(result));
                    }
                    // .Except(other) — set difference
                    "except" => {
                        let other = self.evaluate_expr(&args[0])?;
                        let other_items = other.to_iterable().unwrap_or_default();
                        let other_keys: Vec<String> = other_items.iter().map(|v| v.as_string()).collect();
                        let result: Vec<Value> = items.into_iter()
                            .filter(|v| !other_keys.contains(&v.as_string()))
                            .collect();
                        return Ok(Value::Array(result));
                    }
                    // .Concat(other) — concatenation (allows duplicates)
                    "concat" if !matches!(val, Value::String(_)) => {
                        let other = self.evaluate_expr(&args[0])?;
                        let other_items = other.to_iterable().unwrap_or_default();
                        let mut result = items;
                        result.extend(other_items);
                        return Ok(Value::Array(result));
                    }
                    // .SkipWhile(Function(x) bool)
                    "skipwhile" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let mut skipping = true;
                        let mut result = Vec::new();
                        for item in &items {
                            if skipping && self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                continue;
                            }
                            skipping = false;
                            result.push(item.clone());
                        }
                        return Ok(Value::Array(result));
                    }
                    // .TakeWhile(Function(x) bool)
                    "takewhile" => {
                        let predicate = self.evaluate_expr(&args[0])?;
                        let mut result = Vec::new();
                        for item in &items {
                            if !self.call_lambda(predicate.clone(), &[item.clone()])?.is_truthy() {
                                break;
                            }
                            result.push(item.clone());
                        }
                        return Ok(Value::Array(result));
                    }
                    // .ElementAt(index)
                    "elementat" => {
                        let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                        return items.get(idx).cloned().ok_or_else(|| {
                            RuntimeError::Custom(format!("Index {} out of range", idx))
                        });
                    }
                    // .ElementAtOrDefault(index)
                    "elementatordefault" => {
                        let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                        return Ok(items.get(idx).cloned().unwrap_or(Value::Nothing));
                    }
                    // .DefaultIfEmpty() / .DefaultIfEmpty(defaultValue)
                    "defaultifempty" => {
                        if items.is_empty() {
                            let def = if !args.is_empty() {
                                self.evaluate_expr(&args[0])?
                            } else {
                                Value::Nothing
                            };
                            return Ok(Value::Array(vec![def]));
                        }
                        return Ok(Value::Array(items));
                    }
                    // .SequenceEqual(other)
                    "sequenceequal" => {
                        let other = self.evaluate_expr(&args[0])?;
                        let other_items = other.to_iterable().unwrap_or_default();
                        if items.len() != other_items.len() {
                            return Ok(Value::Boolean(false));
                        }
                        for (a, b) in items.iter().zip(other_items.iter()) {
                            if a.as_string() != b.as_string() {
                                return Ok(Value::Boolean(false));
                            }
                        }
                        return Ok(Value::Boolean(true));
                    }
                    // .ThenBy(Function(x) key) — secondary sort (simplified: just re-sort on this key)
                    "thenby" | "thenbydescending" => {
                        let selector = self.evaluate_expr(&args[0])?;
                        let mut keyed: Vec<(Value, Value)> = Vec::new();
                        for item in &items {
                            let key = self.call_lambda(selector.clone(), &[item.clone()])?;
                            keyed.push((item.clone(), key));
                        }
                        keyed.sort_by(|a, b| compare_values_ordering(&a.1, &b.1));
                        if method_name == "thenbydescending" {
                            keyed.reverse();
                        }
                        return Ok(Value::Array(keyed.into_iter().map(|(v, _)| v).collect()));
                    }
                    _ => {}
                }
            }
        }

        // Handle form methods (extracted to keep call_method stack frame small)
        self.dispatch_form_control_method(obj, &method_name, &object_name, &arg_values)
    }

    /// Dispatch form/control methods like Show, Hide, Close, PerformStep, etc.
    /// Extracted from `call_method` to reduce stack frame size.
    #[inline(never)]
    fn dispatch_form_control_method(&mut self, obj: &Expression, method_name: &str, object_name: &str, arg_values: &[Value]) -> Result<Value, RuntimeError> {
        match method_name {
            "show" => {
                // Show the form
                self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                    object: object_name.to_string(),
                    property: "Visible".to_string(),
                    value: Value::Boolean(true),
                });
                Ok(Value::Nothing)
            }
            "hide" => {
                // Hide the form
                self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                    object: object_name.to_string(),
                    property: "Visible".to_string(),
                    value: Value::Boolean(false),
                });
                Ok(Value::Nothing)
            }
            "move" => {
                // Move form: .Move(left, top, width, height)
                if arg_values.len() >= 2 {
                    self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                        object: object_name.to_string(),
                        property: "Left".to_string(),
                        value: arg_values[0].clone(),
                    });
                    self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                        object: object_name.to_string(),
                        property: "Top".to_string(),
                        value: arg_values[1].clone(),
                    });
                    if arg_values.len() >= 3 {
                        self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                            object: object_name.to_string(),
                            property: "Width".to_string(),
                            value: arg_values[2].clone(),
                        });
                    }
                    if arg_values.len() >= 4 {
                        self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                            object: object_name.to_string(),
                            property: "Height".to_string(),
                            value: arg_values[3].clone(),
                        });
                    }
                }
                Ok(Value::Nothing)
            }
            "navigate" => {
                // Navigate WebBrowser: .Navigate(url)
                if !arg_values.is_empty() {
                    self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                        object: object_name.to_string(),
                        property: "URL".to_string(),
                        value: arg_values[0].clone(),
                    });
                }
                Ok(Value::Nothing)
            }
            "close" | "dispose" => {
                // Form.Close() — fire FormClosing, then FormClosed, then hide
                self.side_effects.push_back(crate::RuntimeSideEffect::FormClose {
                    form_name: object_name.to_string(),
                });
                Ok(Value::Nothing)
            }
            "showdialog" => {
                // Form.ShowDialog() — show form as modal dialog
                self.side_effects.push_back(crate::RuntimeSideEffect::FormShowDialog {
                    form_name: object_name.to_string(),
                });
                // Return DialogResult.OK (1) as default
                Ok(Value::Integer(1))
            }
            "focus" | "select" | "activate" => {
                // Focus/Select — no-op in web renderer but don't error
                Ok(Value::Nothing)
            }
            "bringtofront" | "sendtoback" => {
                Ok(Value::Nothing)
            }
            "refresh" | "invalidate" | "update" | "performlayout" | "suspendlayout" | "resumelayout" => {
                Ok(Value::Nothing)
            }
            "centertoscreen" | "centertoparent" => {
                Ok(Value::Nothing)
            }
            "performstep" => {
                // ProgressBar.PerformStep() — increment Value by Step
                if let Ok(obj_val) = self.evaluate_expr(obj) {
                    if let Value::Object(obj_ref) = &obj_val {
                        let mut b = obj_ref.borrow_mut();
                        let step = b.fields.get("step").and_then(|v| v.as_integer().ok()).unwrap_or(10);
                        let max = b.fields.get("maximum").and_then(|v| v.as_integer().ok()).unwrap_or(100);
                        let cur = b.fields.get("value").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                        let new_val = std::cmp::min(cur + step, max);
                        b.fields.insert("value".to_string(), Value::Integer(new_val));
                    }
                }
                Ok(Value::Nothing)
            }
            "increment" => {
                // NumericUpDown.Increment-like: UpButton/DownButton
                Ok(Value::Nothing)
            }
            "upbutton" => {
                if let Ok(obj_val) = self.evaluate_expr(obj) {
                    if let Value::Object(obj_ref) = &obj_val {
                        let class_name = obj_ref.borrow().class_name.to_lowercase();
                        if class_name == "domainupdown" {
                            // DomainUpDown.UpButton() — move to previous item in Items list
                            let mut b = obj_ref.borrow_mut();
                            let count = b.fields.get("items")
                                .and_then(|v| if let Value::Collection(c) = v { Some(c.borrow().count()) } else { None })
                                .unwrap_or(0i32);
                            let cur = b.fields.get("selectedindex").and_then(|v| v.as_integer().ok()).unwrap_or(-1);
                            if count > 0 {
                                let new_idx = if cur <= 0 { 0i32 } else { cur - 1 };
                                b.fields.insert("selectedindex".to_string(), Value::Integer(new_idx));
                                // Update text to match selected item
                                let item_text = b.fields.get("items")
                                    .and_then(|v| if let Value::Collection(c) = v {
                                        c.borrow().item(new_idx as usize).ok().map(|i| i.as_string())
                                    } else { None })
                                    .unwrap_or_default();
                                b.fields.insert("text".to_string(), Value::String(item_text));
                            }
                        } else {
                            // NumericUpDown.UpButton() — increment Value by Increment
                            let mut b = obj_ref.borrow_mut();
                            let inc = b.fields.get("increment").and_then(|v| v.as_integer().ok()).unwrap_or(1);
                            let max = b.fields.get("maximum").and_then(|v| v.as_integer().ok()).unwrap_or(100);
                            let cur = b.fields.get("value").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                            let new_val = std::cmp::min(cur + inc, max);
                            b.fields.insert("value".to_string(), Value::Integer(new_val));
                        }
                    }
                }
                Ok(Value::Nothing)
            }
            "downbutton" => {
                if let Ok(obj_val) = self.evaluate_expr(obj) {
                    if let Value::Object(obj_ref) = &obj_val {
                        let class_name = obj_ref.borrow().class_name.to_lowercase();
                        if class_name == "domainupdown" {
                            // DomainUpDown.DownButton() — move to next item in Items list
                            let mut b = obj_ref.borrow_mut();
                            let count = b.fields.get("items")
                                .and_then(|v| if let Value::Collection(c) = v { Some(c.borrow().count()) } else { None })
                                .unwrap_or(0i32);
                            let cur = b.fields.get("selectedindex").and_then(|v| v.as_integer().ok()).unwrap_or(-1);
                            if count > 0 {
                                let new_idx = if cur < 0 { 0i32 } else { std::cmp::min(cur + 1, count - 1) };
                                b.fields.insert("selectedindex".to_string(), Value::Integer(new_idx));
                                let item_text = b.fields.get("items")
                                    .and_then(|v| if let Value::Collection(c) = v {
                                        c.borrow().item(new_idx as usize).ok().map(|i| i.as_string())
                                    } else { None })
                                    .unwrap_or_default();
                                b.fields.insert("text".to_string(), Value::String(item_text));
                            }
                        } else {
                            // NumericUpDown.DownButton() — decrement Value by Increment
                            let mut b = obj_ref.borrow_mut();
                            let inc = b.fields.get("increment").and_then(|v| v.as_integer().ok()).unwrap_or(1);
                            let min = b.fields.get("minimum").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                            let cur = b.fields.get("value").and_then(|v| v.as_integer().ok()).unwrap_or(0);
                            let new_val = std::cmp::max(cur - inc, min);
                            b.fields.insert("value".to_string(), Value::Integer(new_val));
                        }
                    }
                }
                Ok(Value::Nothing)
            }
            "setitemchecked" => {
                // CheckedListBox.SetItemChecked(index, checked)
                if arg_values.len() >= 2 {
                    if let Ok(obj_val) = self.evaluate_expr(obj) {
                        if let Value::Object(obj_ref) = &obj_val {
                            let idx = arg_values[0].as_integer().unwrap_or(-1);
                            let checked = match &arg_values[1] {
                                Value::Boolean(b) => *b,
                                Value::Integer(i) => *i != 0,
                                _ => false,
                            };
                            let mut b = obj_ref.borrow_mut();
                            // Store checked states as a simple array in __checked_states
                            let states_key = "__checked_states".to_string();
                            if let Some(Value::Array(states)) = b.fields.get_mut(&states_key) {
                                let idx_u = idx as usize;
                                if idx_u >= states.len() {
                                    states.resize(idx_u + 1, Value::Boolean(false));
                                }
                                states[idx_u] = Value::Boolean(checked);
                            } else {
                                let mut states = vec![];
                                let idx_u = idx as usize;
                                states.resize(idx_u + 1, Value::Boolean(false));
                                states[idx_u] = Value::Boolean(checked);
                                b.fields.insert(states_key, Value::Array(states));
                            }
                        }
                    }
                }
                Ok(Value::Nothing)
            }
            "getitemchecked" => {
                // CheckedListBox.GetItemChecked(index) → Boolean
                if let Some(idx_val) = arg_values.first() {
                    if let Ok(obj_val) = self.evaluate_expr(obj) {
                        if let Value::Object(obj_ref) = &obj_val {
                            let idx = idx_val.as_integer().unwrap_or(-1) as usize;
                            let b = obj_ref.borrow();
                            if let Some(Value::Array(states)) = b.fields.get("__checked_states") {
                                if idx < states.len() {
                                    return Ok(states[idx].clone());
                                }
                            }
                        }
                    }
                }
                Ok(Value::Boolean(false))
            }
            "getitemcheckstate" => {
                // CheckedListBox.GetItemCheckState(index) → CheckState (0=Unchecked, 1=Checked, 2=Indeterminate)
                if let Some(idx_val) = arg_values.first() {
                    if let Ok(obj_val) = self.evaluate_expr(obj) {
                        if let Value::Object(obj_ref) = &obj_val {
                            let idx = idx_val.as_integer().unwrap_or(-1) as usize;
                            let b = obj_ref.borrow();
                            if let Some(Value::Array(states)) = b.fields.get("__checked_states") {
                                if idx < states.len() {
                                    return match &states[idx] {
                                        Value::Boolean(true) => Ok(Value::Integer(1)),
                                        Value::Integer(i) => Ok(Value::Integer(*i)),
                                        _ => Ok(Value::Integer(0)),
                                    };
                                }
                            }
                        }
                    }
                }
                Ok(Value::Integer(0))
            }
            "setitemcheckstate" => {
                // CheckedListBox.SetItemCheckState(index, checkState)
                if arg_values.len() >= 2 {
                    if let Ok(obj_val) = self.evaluate_expr(obj) {
                        if let Value::Object(obj_ref) = &obj_val {
                            let idx = arg_values[0].as_integer().unwrap_or(-1);
                            let state = arg_values[1].as_integer().unwrap_or(0);
                            let mut b = obj_ref.borrow_mut();
                            let states_key = "__checked_states".to_string();
                            if let Some(Value::Array(states)) = b.fields.get_mut(&states_key) {
                                let idx_u = idx as usize;
                                if idx_u >= states.len() {
                                    states.resize(idx_u + 1, Value::Boolean(false));
                                }
                                states[idx_u] = Value::Integer(state);
                            } else {
                                let mut states = vec![];
                                let idx_u = idx as usize;
                                states.resize(idx_u + 1, Value::Integer(0));
                                states[idx_u] = Value::Integer(state);
                                b.fields.insert(states_key, Value::Array(states));
                            }
                            // Also update checkeditems collection
                            drop(b);
                        }
                    }
                }
                Ok(Value::Nothing)
            }
            "selectall" | "clear" | "copy" | "cut" | "paste" | "undo" | "redo" => {
                // TextBox/RichTextBox edit methods — no-op in interpreter
                Ok(Value::Nothing)
            }
            "expandall" | "collapseall" => {
                // TreeView.ExpandAll/CollapseAll — no-op
                Ok(Value::Nothing)
            }
            "beginupdate" | "endupdate" => {
                // ListView/TreeView begin/end update for batch operations — no-op
                Ok(Value::Nothing)
            }
            "getitemat" | "hittest" => {
                // Hit testing — return Nothing
                Ok(Value::Nothing)
            }
            "settooltip" => {
                // ToolTip.SetToolTip(control, text) — store tooltip text for a control
                if arg_values.len() >= 2 {
                    if let Ok(obj_val) = self.evaluate_expr(obj) {
                        if let Value::Object(obj_ref) = &obj_val {
                            let text = arg_values[1].as_string();
                            let ctrl_name = match &arg_values[0] {
                                Value::Object(c) => c.borrow().fields.get("name").map(|v| v.as_string()).unwrap_or_default(),
                                Value::String(s) => s.clone(),
                                _ => String::new(),
                            };
                            let mut b = obj_ref.borrow_mut();
                            let tooltips = b.fields.entry("__tooltips".to_string()).or_insert_with(|| Value::Nothing);
                            if let Value::Object(map_ref) = tooltips {
                                map_ref.borrow_mut().fields.insert(ctrl_name.to_lowercase(), Value::String(text));
                            } else {
                                let mut fields = std::collections::HashMap::new();
                                fields.insert(ctrl_name.to_lowercase(), Value::String(text));
                                let map_obj = crate::value::ObjectData { class_name: "__TooltipMap".to_string(), fields };
                                *tooltips = Value::Object(std::rc::Rc::new(std::cell::RefCell::new(map_obj)));
                            }
                        }
                    }
                }
                Ok(Value::Nothing)
            }
            "gettooltip" => {
                // ToolTip.GetToolTip(control) — retrieve tooltip text for a control
                if !arg_values.is_empty() {
                    if let Ok(obj_val) = self.evaluate_expr(obj) {
                        if let Value::Object(obj_ref) = &obj_val {
                            let ctrl_name = match &arg_values[0] {
                                Value::Object(c) => c.borrow().fields.get("name").map(|v| v.as_string()).unwrap_or_default(),
                                Value::String(s) => s.clone(),
                                _ => String::new(),
                            };
                            let b = obj_ref.borrow();
                            if let Some(Value::Object(map_ref)) = b.fields.get("__tooltips") {
                                let map = map_ref.borrow();
                                if let Some(val) = map.fields.get(&ctrl_name.to_lowercase()) {
                                    return Ok(val.clone());
                                }
                            }
                        }
                    }
                }
                Ok(Value::String(String::new()))
            }
            "removeall" => {
                // ToolTip.RemoveAll() — clear all tooltip associations
                if let Ok(obj_val) = self.evaluate_expr(obj) {
                    if let Value::Object(obj_ref) = &obj_val {
                        let mut b = obj_ref.borrow_mut();
                        b.fields.insert("__tooltips".to_string(), Value::Nothing);
                    }
                }
                Ok(Value::Nothing)
            }
            "setrange" => {
                // MonthCalendar.SetSelectionRange(start, end) or scrollbar range
                if arg_values.len() >= 2 {
                    if let Ok(obj_val) = self.evaluate_expr(obj) {
                        if let Value::Object(obj_ref) = &obj_val {
                            let mut b = obj_ref.borrow_mut();
                            b.fields.insert("selectionstart".to_string(), arg_values[0].clone());
                            b.fields.insert("selectionend".to_string(), arg_values[1].clone());
                        }
                    }
                }
                Ok(Value::Nothing)
            }
            "setselectionrange" => {
                // MonthCalendar.SetSelectionRange(date1, date2)
                if arg_values.len() >= 2 {
                    if let Ok(obj_val) = self.evaluate_expr(obj) {
                        if let Value::Object(obj_ref) = &obj_val {
                            let mut b = obj_ref.borrow_mut();
                            b.fields.insert("selectionstart".to_string(), arg_values[0].clone());
                            b.fields.insert("selectionend".to_string(), arg_values[1].clone());
                        }
                    }
                }
                Ok(Value::Nothing)
            }
            "setcalendarmonthlybolddates" | "setcalendarannuallybolddates" | "removeallboldeddates" |
            "removeallannuallybolddates" | "removeallmonthlybolddates" | "addboldeddate" |
            "addannuallyboldeddate" | "addmonthlyboldeddate" | "removeboldeddate" |
            "removeannuallyboldeddate" | "removemonthlyboldeddate" | "updateboldeddates" => {
                // MonthCalendar bold date management — no-op for now
                Ok(Value::Nothing)
            }
            _ => Err(RuntimeError::Custom(format!("Unknown method: {}", method_name))),
        }
    }

    /// Try to find and call an extension method for the given object and method name.
    /// Extension methods are Subs/Functions marked with <Extension()> whose first parameter
    /// receives the object instance.
    fn try_extension_method(&mut self, obj: &Expression, method: &Identifier, args: &[Expression]) -> Result<Value, RuntimeError> {
        let method_lower = method.as_str().to_lowercase();

        // Evaluate the object (the "self" for the extension method)
        let obj_val = self.evaluate_expr(obj)?;

        // Evaluate the remaining arguments
        let mut arg_vals: Vec<Value> = Vec::with_capacity(args.len() + 1);
        arg_vals.push(obj_val); // First arg = the object
        for arg in args {
            arg_vals.push(self.evaluate_expr(arg)?);
        }

        // Search extension functions
        for (_key, func) in &self.functions {
            if func.is_extension && func.name.as_str().to_lowercase() == method_lower {
                let func_clone = func.clone();
                return self.call_user_function(&func_clone, &arg_vals, None);
            }
        }

        // Search extension subs
        for (_key, sub) in &self.subs {
            if sub.is_extension && sub.name.as_str().to_lowercase() == method_lower {
                let sub_clone = sub.clone();
                self.call_user_sub(&sub_clone, &arg_vals, None)?;
                return Ok(Value::Nothing);
            }
        }

        Err(RuntimeError::Custom(format!("No extension method found: {}", method.as_str())))
    }

    // Generic helper for subs
    fn call_user_sub_impl(&mut self, sub: &SubDecl, args: Option<&[Value]>, arg_exprs: Option<&[Expression]>, context: Option<Rc<RefCell<ObjectData>>>) -> Result<Value, RuntimeError> {
        // Push new scope
        self.env.push_scope();

        // Save previous object context and set new one
        let prev_object = self.current_object.take();
        self.current_object = context;
        let prev_procedure = self.current_procedure.take();
        self.current_procedure = Some(sub.name.as_str().to_string());

        let mut byref_writebacks = Vec::new();

        // Bind parameters
        for (i, param) in sub.parameters.iter().enumerate() {
            let mut val = Value::Nothing;
            
            // ParamArray: last parameter collects remaining args into an array
            if param.is_param_array {
                if let Some(values) = args {
                    let remaining: Vec<Value> = values[i..].to_vec();
                    val = Value::Array(remaining);
                } else if let Some(exprs) = arg_exprs {
                    let mut remaining = Vec::new();
                    for expr in &exprs[i..] {
                        remaining.push(self.evaluate_expr(expr)?);
                    }
                    val = Value::Array(remaining);
                }
                self.env.define(param.name.as_str(), val);
                break; // ParamArray must be last parameter
            }

            if let Some(values) = args {
                if i < values.len() {
                    val = values[i].clone();
                } else if param.is_optional {
                    if let Some(ref default) = param.default_value {
                        val = self.evaluate_expr(default)?;
                    }
                }
            } else if let Some(exprs) = arg_exprs {
                if i < exprs.len() {
                    val = self.evaluate_expr(&exprs[i])?;
                    let is_byref = param.pass_type == vybe_parser::ast::decl::ParameterPassType::ByRef;
                    if is_byref {
                        // Check if argument is a variable we can write back to
                        match &exprs[i] {
                            Expression::Variable(name) => {
                                byref_writebacks.push((name.as_str().to_string(), param.name.as_str().to_string()));
                            }
                            _ => {}
                        }
                    }
                }
            }

            self.env.define(param.name.as_str(), val);
        }

        // Execute body with GoTo and On Error support
        let result = self.execute_body_with_goto(&sub.body);
        match result {
            Err(RuntimeError::Exit(ExitType::Sub)) => {}
            Err(e) => {
                self.env.pop_scope();
                self.current_object = prev_object;
                self.current_procedure = prev_procedure;
                return Err(e);
            }
            Ok(_) => {}
        }

        // Perform ByRef writebacks (capture values before popping scope)
        let mut final_writebacks = Vec::new();
        if !byref_writebacks.is_empty() {
             for (caller_var, param_name) in byref_writebacks {
                 if let Ok(new_val) = self.env.get(&param_name) {
                     final_writebacks.push((caller_var, new_val));
                 }
             }
        }

        // Persist static local variables before popping scope
        self.persist_static_locals();

        // Pop scope
        self.env.pop_scope();
        self.current_object = prev_object;
        self.current_procedure = prev_procedure;
        
        // Apply writebacks in restored scope
        for (var_name, val) in final_writebacks {
            let _ = self.env.set(&var_name, val);
        }

        Ok(Value::Nothing)
    }

    fn call_user_sub(&mut self, sub: &SubDecl, args: &[Value], context: Option<Rc<RefCell<ObjectData>>>) -> Result<Value, RuntimeError> {
        self.call_user_sub_impl(sub, Some(args), None, context)
    }

    fn call_user_sub_exprs(&mut self, sub: &SubDecl, args: &[Expression], context: Option<Rc<RefCell<ObjectData>>>) -> Result<Value, RuntimeError> {
        self.call_user_sub_impl(sub, None, Some(args), context)
    }

    // Generic helper for functions
    fn call_user_function_impl(&mut self, func: &FunctionDecl, args: Option<&[Value]>, arg_exprs: Option<&[Expression]>, context: Option<Rc<RefCell<ObjectData>>>) -> Result<Value, RuntimeError> {
        self.env.push_scope();
        let prev_object = self.current_object.take();
        self.current_object = context;
        let prev_procedure = self.current_procedure.take();
        self.current_procedure = Some(func.name.as_str().to_string());

        let mut byref_writebacks = Vec::new();

        for (i, param) in func.parameters.iter().enumerate() {
            let mut val = Value::Nothing;

            // ParamArray: last parameter collects remaining args into an array
            if param.is_param_array {
                if let Some(values) = args {
                    let remaining: Vec<Value> = values[i..].to_vec();
                    val = Value::Array(remaining);
                } else if let Some(exprs) = arg_exprs {
                    let mut remaining = Vec::new();
                    for expr in &exprs[i..] {
                        remaining.push(self.evaluate_expr(expr)?);
                    }
                    val = Value::Array(remaining);
                }
                self.env.define(param.name.as_str(), val);
                break; // ParamArray must be last parameter
            }

            if let Some(values) = args {
                if i < values.len() {
                    val = values[i].clone();
                } else if param.is_optional {
                    if let Some(ref default) = param.default_value {
                        val = self.evaluate_expr(default)?;
                    }
                }
            } else if let Some(exprs) = arg_exprs {
                if i < exprs.len() {
                    val = self.evaluate_expr(&exprs[i])?;
                    let is_byref = param.pass_type == vybe_parser::ast::decl::ParameterPassType::ByRef;
                    if is_byref {
                        match &exprs[i] {
                            Expression::Variable(name) => {
                                byref_writebacks.push((name.as_str().to_string(), param.name.as_str().to_string()));
                            }
                            _ => {}
                        }
                    }
                } else if param.is_optional {
                    if let Some(ref default) = param.default_value {
                        val = self.evaluate_expr(default)?;
                    }
                }
            }
            self.env.define(param.name.as_str(), val);
        }

        self.env.define(func.name.as_str(), Value::Nothing);
        let mut explicit_return: Option<Value> = None;

        let mut result = Value::Nothing;
        // Execute body with GoTo and On Error support
        let body = func.body.clone();
        let exec_result = self.execute_body_with_goto(&body);
        match exec_result {
            Err(RuntimeError::Exit(ExitType::Function)) => {}
            Err(RuntimeError::Return(val)) => {
                if let Some(v) = val {
                    explicit_return = Some(v);
                }
            }
            Err(e) => {
                self.env.pop_scope();
                self.current_object = prev_object;
                self.current_procedure = prev_procedure;
                return Err(e);
            }
            Ok(_) => {}
        }

        if let Some(ret) = explicit_return {
            result = ret;
        } else if let Ok(val) = self.env.get(func.name.as_str()) {
             result = val;
        }

        // Capture byref values before popping
        let mut final_writebacks = Vec::new();
        if !byref_writebacks.is_empty() {
             for (caller_var, param_name) in byref_writebacks {
                 if let Ok(new_val) = self.env.get(&param_name) {
                     final_writebacks.push((caller_var, new_val));
                 }
             }
        }

        // Persist static local variables before popping scope
        self.persist_static_locals();

        self.env.pop_scope();
        self.current_object = prev_object;
        self.current_procedure = prev_procedure;

        // Writeback
        for (var_name, val) in final_writebacks {
            let _ = self.env.set(&var_name, val);
        }

        Ok(result)
    }

    fn call_user_function(&mut self, func: &FunctionDecl, args: &[Value], context: Option<Rc<RefCell<ObjectData>>>) -> Result<Value, RuntimeError> {
         self.call_user_function_impl(func, Some(args), None, context)
    }

    fn call_user_function_exprs(&mut self, func: &FunctionDecl, args: &[Expression], context: Option<Rc<RefCell<ObjectData>>>) -> Result<Value, RuntimeError> {
         self.call_user_function_impl(func, None, Some(args), context)
    }

    pub fn call_event_handler(&mut self, handler_name: &str, args: &[Value]) -> Result<(), RuntimeError> {
        let handler_name_lower = handler_name.to_lowercase();

        // Try to find the sub with or without module prefix
        // First, try all qualified names to find which module it belongs to
        let mut found_sub = None;
        let mut module_name = None;

        for (key, sub) in &self.subs {
            if key == &handler_name_lower {
                // Found unqualified
                found_sub = Some(sub.clone());
                break;
            } else if key.ends_with(&format!(".{}", handler_name_lower)) {
                // Found qualified (module.handler)
                let parts: Vec<&str> = key.split('.').collect();
                if parts.len() == 2 {
                    module_name = Some(parts[0].to_string());
                    found_sub = Some(sub.clone());
                    break;
                }
            }
        }

        if let Some(sub) = found_sub {
            // Set module context if we found one
            let prev_module = self.current_module.clone();
            if let Some(module) = module_name {
                self.current_module = Some(module);
            }
            let result = self.call_user_sub(&sub, args, None);
            if let Err(e) = result {
                // Restore previous module
                self.current_module = prev_module;
                return Err(e);
            }

            // Restore previous module
            self.current_module = prev_module;
            Ok(())
        } else {
            Err(RuntimeError::UndefinedFunction(handler_name.to_string()))
        }
    }

    pub fn call_instance_method(&mut self, instance_name: &str, method_name: &str, args: &[Value]) -> Result<(), RuntimeError> {
        let instance_val = self.env.get(instance_name)?;
        if let Value::Object(obj_ref) = instance_val {
            self.call_method_on_object(&obj_ref, method_name, args)
        } else {
            Err(RuntimeError::TypeError { expected: "Object".to_string(), got: format!("{:?}", instance_val) })
        }
    }

    /// Call a method on an object reference directly (no environment lookup needed).
    pub fn call_method_on_object(
        &mut self,
        obj_ref: &std::rc::Rc<std::cell::RefCell<crate::value::ObjectData>>,
        method_name: &str,
        args: &[Value],
    ) -> Result<(), RuntimeError> {
        let class_name = obj_ref.borrow().class_name.clone();
        if let Some(method) = self.find_method(&class_name, method_name) {
            match method {
                vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                    self.call_user_sub(&s, args, Some(obj_ref.clone()))?;
                    Ok(())
                }
                vybe_parser::ast::decl::MethodDecl::Function(f) => {
                    let _ = self.call_user_function(&f, args, Some(obj_ref.clone()))?;
                    Ok(())
                }
            }
        } else {
            Err(RuntimeError::UndefinedFunction(method_name.to_string()))
        }
    }

    /// Create an instance of a registered class and optionally call Sub New / InitializeComponent.
    /// This is the Rust-side equivalent of `New ClassName()` without needing to parse VB code.
    /// Returns the Rc-wrapped object so the caller can store it in the environment.
    pub fn create_class_instance(&mut self, class_name: &str) -> Result<std::rc::Rc<std::cell::RefCell<crate::value::ObjectData>>, RuntimeError> {
        let resolved_key = self.resolve_class_key(class_name);
        let class_decl = resolved_key
            .as_ref()
            .and_then(|k| self.classes.get(k).cloned())
            .ok_or_else(|| RuntimeError::Custom(format!("Class '{}' not found", class_name)))?;

        // Enforce MustInherit
        if class_decl.is_must_inherit {
            return Err(RuntimeError::Custom(format!(
                "Cannot create an instance of MustInherit class '{}'",
                class_decl.name.as_str()
            )));
        }

        let fields = self.collect_fields(class_name);
        let obj_data = crate::value::ObjectData {
            class_name: class_decl.name.as_str().to_string(),
            fields,
        };
        let obj_ref = std::rc::Rc::new(std::cell::RefCell::new(obj_data));

        // Call Sub New if it exists (which typically calls InitializeComponent)
        let new_method = self.find_method(class_name, "new");
        if let Some(method) = new_method {
            // Auto-call base Sub New first if derived doesn't explicitly do it
            let method_body = match &method {
                vybe_parser::ast::decl::MethodDecl::Sub(s) => &s.body,
                vybe_parser::ast::decl::MethodDecl::Function(f) => &f.body,
            };
            if !body_contains_mybase_new(method_body) {
                if let Some(base_new) = self.find_method_in_base(class_name, "new") {
                    match base_new {
                        vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                            let _ = self.call_user_sub(&s, &[], Some(obj_ref.clone()));
                        }
                        vybe_parser::ast::decl::MethodDecl::Function(f) => {
                            let _ = self.call_user_function(&f, &[], Some(obj_ref.clone()));
                        }
                    }
                }
            }

            match method {
                vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                    self.call_user_sub(&s, &[], Some(obj_ref.clone()))?;
                }
                vybe_parser::ast::decl::MethodDecl::Function(f) => {
                    let _ = self.call_user_function(&f, &[], Some(obj_ref.clone()))?;
                }
            }
        } else {
            // No Sub New: auto-call InitializeComponent for form classes
            let inherits_form = class_decl.inherits.as_ref().map_or(false, |t| {
                match t {
                    vybe_parser::VBType::Custom(n) => n.to_lowercase().contains("form"),
                    _ => false,
                }
            });
            if inherits_form {
                if let Some(init_method) = self.find_method(class_name, "InitializeComponent") {
                    match init_method {
                        vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                            if let Err(e) = self.call_user_sub(&s, &[], Some(obj_ref.clone())) {
                                eprintln!("[InitializeComponent error] {}", e);
                            }
                        }
                        vybe_parser::ast::decl::MethodDecl::Function(f) => {
                            let _ = self.call_user_function(&f, &[], Some(obj_ref.clone()));
                        }
                    }
                }
            }
        }

        Ok(obj_ref)
    }

    /// Find a class method that has a Handles clause matching the given control.event pattern.
    /// For example, find_handles_method("form1", "btn0", "Click") finds a method with `Handles btn0.Click`.
    /// For Me.Load, use control_name="Me" and event_name="Load".
    /// Find all methods that handle the given control.event.
    /// This includes methods with `Handles` clauses and handlers registered via `AddHandler`.
    pub fn get_event_handlers(&self, class_name: &str, control_name: &str, event_name: &str) -> Vec<String> {
        let mut handlers = Vec::new();
        let key = class_name.to_lowercase();
        let target = format!("{}.{}", control_name, event_name).to_lowercase();

        // 1. Find static `Handles` clauses
        if let Some(cls) = self.classes.get(&key) {
            for method in &cls.methods {
                if let vybe_parser::ast::decl::MethodDecl::Sub(s) = method {
                    if let Some(handles_list) = &s.handles {
                        for h in handles_list {
                            if h.to_lowercase() == target {
                                handlers.push(s.name.as_str().to_string());
                            }
                        }
                    }
                }
            }
        }

        // 2. Find dynamic `AddHandler` registrations
        // Note: Runtime events are stored as "control_event", so we query "btn1" and "Click" (EventType)
        // However, EventSystem stores keys as "control_event".
        if let Some(event_type) = vybe_forms::EventType::from_name(event_name) {
            if let Some(dynamic_handlers) = self.events.get_handlers(control_name, &event_type) {
                handlers.extend_from_slice(dynamic_handlers);
            }
        }

        handlers
    }

    /// Get all Handles clause mappings for a class as (control, event) -> method_name.
    pub fn get_handles_map(&self, class_name: &str) -> HashMap<(String, String), String> {
        let mut map = HashMap::new();
        let key = class_name.to_lowercase();
        if let Some(cls) = self.classes.get(&key) {
            for method in &cls.methods {
                if let vybe_parser::ast::decl::MethodDecl::Sub(s) = method {
                    if let Some(handles) = &s.handles {
                        for h in handles {
                            let parts: Vec<&str> = h.splitn(2, '.').collect();
                            if parts.len() == 2 {
                                map.insert(
                                    (parts[0].to_lowercase(), parts[1].to_lowercase()),
                                    s.name.as_str().to_string(),
                                );
                            }
                        }
                    }
                }
            }
        }
        map
    }

    pub fn trigger_event(&mut self, control_name: &str, event_type: vybe_forms::EventType, index: Option<i32>) -> Result<(), RuntimeError> {
        let handlers = if let Some(h) = self.events.get_handlers(control_name, &event_type) {
            h.clone()
        } else {
            Vec::new()
        };
        
        if handlers.is_empty() { return Ok(()); }
        
        // Prepare args once
        let args: Vec<Value> = if let Some(idx) = index {
             vec![Value::Integer(idx)]
        } else {
             self.make_event_handler_args(control_name, event_type.as_str())
        };
        
        for handler_name in handlers {
            if self.subs.contains_key(&handler_name.to_lowercase()) {
                self.call_event_handler(&handler_name, &args)?;
            }
        }
        Ok(())
    }

    fn execute_block(&mut self, stmts: &[Statement]) -> Result<(), RuntimeError> {
        for stmt in stmts {
            self.execute(stmt)?;
        }
        Ok(())
    }

    /// Refresh all controls bound to a BindingSource after the position has changed.
    /// Reads the __bindings array (entries: "controlName|propertyName|dataMember")
    /// and pushes PropertyChange side effects for each.
    #[allow(dead_code)]
    fn refresh_bindings(
        &mut self,
        bs_ref: &std::rc::Rc<std::cell::RefCell<crate::value::ObjectData>>,
        datasource: &Value,
        position: i32,
    ) {
        let bindings: Vec<String> = bs_ref.borrow()
            .fields.get("__bindings")
            .and_then(|v| if let Value::Array(arr) = v {
                Some(arr.iter().filter_map(|v| if let Value::String(s) = v { Some(s.clone()) } else { None }).collect())
            } else { None })
            .unwrap_or_default();

        if bindings.is_empty() { return; }

        let row = self.binding_source_get_row(datasource, position);
        if let Value::Object(row_ref) = &row {
            for entry in &bindings {
                let parts: Vec<&str> = entry.split('|').collect();
                if parts.len() >= 3 {
                    let ctrl_name = parts[0];
                    let prop_name = parts[1];
                    let data_member = parts[2].to_lowercase();
                    let cell_val = row_ref.borrow().fields.get(&data_member)
                        .cloned().unwrap_or(Value::String(String::new()));
                    self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                        object: ctrl_name.to_string(),
                        property: prop_name.to_string(),
                        value: cell_val,
                    });
                }
            }
        }
    }

    /// Like `refresh_bindings` but applies filter/sort from the BindingSource.
    /// `position` is the index into the filtered/sorted results.
    pub fn refresh_bindings_filtered(
        &mut self,
        bs_ref: &std::rc::Rc<std::cell::RefCell<crate::value::ObjectData>>,
        datasource: &Value,
        position: i32,
    ) {
        let bindings: Vec<String> = bs_ref.borrow()
            .fields.get("__bindings")
            .and_then(|v| if let Value::Array(arr) = v {
                Some(arr.iter().filter_map(|v| if let Value::String(s) = v { Some(s.clone()) } else { None }).collect())
            } else { None })
            .unwrap_or_default();

        if bindings.is_empty() { return; }

        let filter = bs_ref.borrow().fields.get("filter")
            .map(|v| v.as_string()).unwrap_or_default();
        let sort = bs_ref.borrow().fields.get("sort")
            .map(|v| v.as_string()).unwrap_or_default();

        // Efficiently get the target row value
        let row_val = if filter.is_empty() && sort.is_empty() {
            // Direct access: avoids re-creating all rows as ObjectData
            self.binding_source_get_row(datasource, position)
        } else {
            // Filter/Sort active: unfortunately must process all rows
            let all_rows = self.get_all_data_rows(datasource);
            let filtered = Self::apply_filter_sort(&all_rows, &filter, &sort);
            filtered.get(position as usize).cloned().unwrap_or(Value::Nothing)
        };

        if let Value::Object(r) = row_val {
            for entry in &bindings {
                let parts: Vec<&str> = entry.split('|').collect();
                if parts.len() >= 3 {
                    let ctrl_name = parts[0];
                    let prop_name = parts[1];
                    let data_member = parts[2].to_lowercase();

                    let cell_val = r.borrow().fields.get(&data_member)
                        .cloned().unwrap_or(Value::String(String::new()));

                    self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                        object: ctrl_name.to_string(),
                        property: prop_name.to_string(),
                        value: cell_val,
                    });
                }
            }
        }
    }

    /// Get the row count from a BindingSource, applying its filter.
    pub fn binding_source_row_count_filtered(&self, datasource: &Value) -> i32 {
        if let Value::Object(obj_ref) = datasource {
            let obj = obj_ref.borrow();
            let dt_type = obj.fields.get("__type")
                .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                .unwrap_or_default();
            if dt_type == "BindingSource" {
                let inner_ds = obj.fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                let filter = obj.fields.get("filter").map(|v| v.as_string()).unwrap_or_default();
                let sort = obj.fields.get("sort").map(|v| v.as_string()).unwrap_or_default();
                let dm = obj.fields.get("datamember").map(|v| v.as_string()).unwrap_or_default();
                drop(obj);
                Self::inject_select_from_data_member(&inner_ds, &dm);
                if filter.is_empty() && sort.is_empty() {
                    return self.binding_source_row_count(&inner_ds);
                }
                let all_rows = self.get_all_data_rows(&inner_ds);
                let filtered = Self::apply_filter_sort(&all_rows, &filter, &sort);
                return filtered.len() as i32;
            }
        }
        self.binding_source_row_count(datasource)
    }

    /// Get all DataRows from a datasource as Value::Object(DataRow) objects.
    fn get_all_data_rows(&self, datasource: &Value) -> Vec<Value> {
        let count = self.binding_source_row_count(datasource);
        let mut rows = Vec::new();
        for i in 0..count {
            let row = self.binding_source_get_row(datasource, i);
            if !matches!(row, Value::Nothing) {
                rows.push(row);
            }
        }
        rows
    }

    /// Get columns+rows for DataGridView, applying BindingSource filter/sort.
    pub fn get_datasource_table_data_filtered(&self, datasource: &Value) -> (Vec<String>, Vec<Vec<String>>) {
        if let Value::Object(obj_ref) = datasource {
            let obj = obj_ref.borrow();
            let dt_type = obj.fields.get("__type")
                .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                .unwrap_or_default();
            if dt_type == "BindingSource" {
                let inner_ds = obj.fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                let filter = obj.fields.get("filter").map(|v| v.as_string()).unwrap_or_default();
                let sort = obj.fields.get("sort").map(|v| v.as_string()).unwrap_or_default();
                let dm = obj.fields.get("datamember").map(|v| v.as_string()).unwrap_or_default();
                drop(obj);
                Self::inject_select_from_data_member(&inner_ds, &dm);

                if filter.is_empty() && sort.is_empty() {
                    return self.get_datasource_table_data(&inner_ds);
                }

                // Get full data, filter, sort, return
                let (columns, all_rows) = self.get_datasource_table_data(&inner_ds);
                if columns.is_empty() { return (columns, all_rows); }

                let col_lower: Vec<String> = columns.iter().map(|c| c.to_lowercase()).collect();
                let filtered_rows = Self::apply_filter_sort_raw(&columns, &col_lower, &all_rows, &filter, &sort);
                return (columns, filtered_rows);
            }
        }
        self.get_datasource_table_data(datasource)
    }

    /// Apply a VB.NET-style Filter string and Sort string to a list of DataRow objects.
    /// Filter example: "Name = 'John'" or "Age > 30"
    /// Sort example: "Name ASC" or "Age DESC, Name ASC"
    #[inline(never)]
    fn apply_filter_sort(rows: &[Value], filter: &str, sort: &str) -> Vec<Value> {
        let mut result: Vec<Value> = if filter.trim().is_empty() {
            rows.to_vec()
        } else {
            rows.iter().filter(|row| {
                if let Value::Object(r) = row {
                    Self::row_matches_filter(&r.borrow().fields, filter)
                } else {
                    true
                }
            }).cloned().collect()
        };

        if !sort.trim().is_empty() {
            let sort_specs = Self::parse_sort_spec(sort);
            result.sort_by(|a, b| {
                for (col, ascending) in &sort_specs {
                    let va = if let Value::Object(r) = a { r.borrow().fields.get(col).cloned().unwrap_or(Value::Nothing) } else { Value::Nothing };
                    let vb = if let Value::Object(r) = b { r.borrow().fields.get(col).cloned().unwrap_or(Value::Nothing) } else { Value::Nothing };
                    let cmp = va.as_string().cmp(&vb.as_string());
                    let cmp = if *ascending { cmp } else { cmp.reverse() };
                    if cmp != std::cmp::Ordering::Equal { return cmp; }
                }
                std::cmp::Ordering::Equal
            });
        }
        result
    }

    /// Apply filter/sort to raw string rows (Vec<Vec<String>>) for DataGridView rendering.
    #[inline(never)]
    fn apply_filter_sort_raw(_columns: &[String], col_lower: &[String], rows: &[Vec<String>], filter: &str, sort: &str) -> Vec<Vec<String>> {
        let mut result: Vec<Vec<String>> = if filter.trim().is_empty() {
            rows.to_vec()
        } else {
            rows.iter().filter(|row| {
                // Build a temporary fields map for the filter check
                let mut fields = std::collections::HashMap::new();
                for (i, col) in col_lower.iter().enumerate() {
                    if let Some(val) = row.get(i) {
                        fields.insert(col.clone(), Value::String(val.clone()));
                    }
                }
                Self::row_matches_filter(&fields, filter)
            }).cloned().collect()
        };

        if !sort.trim().is_empty() {
            let sort_specs = Self::parse_sort_spec(sort);
            result.sort_by(|a, b| {
                for (col, ascending) in &sort_specs {
                    let idx = col_lower.iter().position(|c| c == col).unwrap_or(usize::MAX);
                    let va = a.get(idx).cloned().unwrap_or_default();
                    let vb = b.get(idx).cloned().unwrap_or_default();
                    // Try numeric comparison first
                    let cmp = match (va.parse::<f64>(), vb.parse::<f64>()) {
                        (Ok(na), Ok(nb)) => na.partial_cmp(&nb).unwrap_or(std::cmp::Ordering::Equal),
                        _ => va.cmp(&vb),
                    };
                    let cmp = if *ascending { cmp } else { cmp.reverse() };
                    if cmp != std::cmp::Ordering::Equal { return cmp; }
                }
                std::cmp::Ordering::Equal
            });
        }
        result
    }

    /// Parse sort spec like "Name ASC, Age DESC" → vec![("name", true), ("age", false)]
    fn parse_sort_spec(sort: &str) -> Vec<(String, bool)> {
        sort.split(',')
            .filter_map(|part| {
                let tokens: Vec<&str> = part.trim().split_whitespace().collect();
                if tokens.is_empty() { return None; }
                let col = tokens[0].to_lowercase();
                let ascending = tokens.get(1)
                    .map(|d| !d.eq_ignore_ascii_case("desc"))
                    .unwrap_or(true);
                Some((col, ascending))
            })
            .collect()
    }

    /// Check if a row's fields match a simple VB.NET filter expression.
    /// Supports: col = 'val', col <> 'val', col > val, col < val, col >= val, col <= val,
    /// col LIKE 'pattern%', and AND/OR combinators.
    fn row_matches_filter(fields: &std::collections::HashMap<String, Value>, filter: &str) -> bool {
        let filter = filter.trim();
        if filter.is_empty() { return true; }

        // Handle AND/OR (simple split — doesn't handle nested parens)
        let filter_upper = filter.to_uppercase();
        if let Some(pos) = Self::find_logical_op(&filter_upper, " AND ") {
            let left = &filter[..pos];
            let right = &filter[pos + 5..];
            return Self::row_matches_filter(fields, left) && Self::row_matches_filter(fields, right);
        }
        if let Some(pos) = Self::find_logical_op(&filter_upper, " OR ") {
            let left = &filter[..pos];
            let right = &filter[pos + 4..];
            return Self::row_matches_filter(fields, left) || Self::row_matches_filter(fields, right);
        }

        // Parse single condition: column op value
        // Operators: =, <>, !=, >, <, >=, <=, LIKE
        let ops = ["<>", "!=", ">=", "<=", ">", "<", "="];
        for op in &ops {
            if let Some(idx) = filter.find(op) {
                let col = filter[..idx].trim().to_lowercase();
                let val_str = filter[idx + op.len()..].trim();
                let val_str = Self::unquote(val_str);
                let field_val = fields.get(&col).map(|v| v.as_string()).unwrap_or_default();
                return Self::compare_filter_values(&field_val, op, &val_str);
            }
        }
        // LIKE operator
        if let Some(idx) = filter_upper.find(" LIKE ") {
            let col = filter[..idx].trim().to_lowercase();
            let pattern = filter[idx + 6..].trim();
            let pattern = Self::unquote(pattern);
            let field_val = fields.get(&col).map(|v| v.as_string()).unwrap_or_default();
            return Self::like_match(&field_val, &pattern);
        }
        true // unrecognized filter → pass through
    }

    /// Find position of a logical operator not inside quotes.
    fn find_logical_op(s: &str, op: &str) -> Option<usize> {
        let mut in_quote = false;
        let bytes = s.as_bytes();
        for i in 0..s.len() {
            if bytes[i] == b'\'' { in_quote = !in_quote; }
            if !in_quote && s[i..].starts_with(op) {
                return Some(i);
            }
        }
        None
    }

    /// Remove surrounding quotes from a filter value.
    fn unquote(s: &str) -> String {
        let s = s.trim();
        if (s.starts_with('\'') && s.ends_with('\'')) || (s.starts_with('"') && s.ends_with('"')) {
            s[1..s.len()-1].to_string()
        } else {
            s.to_string()
        }
    }

    /// Compare a field value against a filter value using the given operator.
    fn compare_filter_values(field_val: &str, op: &str, filter_val: &str) -> bool {
        // Try numeric comparison
        if let (Ok(fv), Ok(rv)) = (field_val.parse::<f64>(), filter_val.parse::<f64>()) {
            return match op {
                "=" => (fv - rv).abs() < f64::EPSILON,
                "<>" | "!=" => (fv - rv).abs() >= f64::EPSILON,
                ">" => fv > rv,
                "<" => fv < rv,
                ">=" => fv >= rv,
                "<=" => fv <= rv,
                _ => false,
            };
        }
        // String comparison (case-insensitive)
        let fv = field_val.to_lowercase();
        let rv = filter_val.to_lowercase();
        match op {
            "=" => fv == rv,
            "<>" | "!=" => fv != rv,
            ">" => fv > rv,
            "<" => fv < rv,
            ">=" => fv >= rv,
            "<=" => fv <= rv,
            _ => false,
        }
    }

    /// Simple LIKE pattern matching with % wildcards.
    fn like_match(value: &str, pattern: &str) -> bool {
        let val = value.to_lowercase();
        let pat = pattern.to_lowercase();
        if pat.starts_with('%') && pat.ends_with('%') && pat.len() > 1 {
            val.contains(&pat[1..pat.len()-1])
        } else if pat.starts_with('%') {
            val.ends_with(&pat[1..])
        } else if pat.ends_with('%') {
            val.starts_with(&pat[..pat.len()-1])
        } else {
            val == pat
        }
    }

    /// If `ds` is a DataAdapter that has no SelectCommand yet and `data_member`
    /// is non-empty, inject `SELECT * FROM <data_member>` as the selectcommandtext.
    fn inject_select_from_data_member(ds: &Value, data_member: &str) {
        if data_member.is_empty() { return; }
        if let Value::Object(da_ref) = ds {
            let da = da_ref.borrow();
            let is_da = da.fields.get("__type")
                .map(|v| v.as_string() == "DataAdapter")
                .unwrap_or(false);
            if !is_da { return; }
            let has_sql = da.fields.get("selectcommandtext")
                .or(da.fields.get("selectcommand"))
                .map(|v| !v.as_string().is_empty())
                .unwrap_or(false);
            drop(da);
            if !has_sql {
                da_ref.borrow_mut().fields.insert(
                    "selectcommandtext".to_string(),
                    Value::String(format!("SELECT * FROM {}", data_member)),
                );
            }
        }
    }

    /// Auto-fill a DataAdapter: open a connection using its ConnectionString,
    /// execute its SelectCommand, and store the resulting recordset ID on the adapter.
    /// This is called lazily the first time a BindingSource (or control) tries
    /// to read data through a DataAdapter that hasn't been filled yet.
    fn auto_fill_data_adapter(da_ref: &std::rc::Rc<std::cell::RefCell<crate::value::ObjectData>>) {
        // Already filled?
        let already = da_ref.borrow().fields.get("__rs_id")
            .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
            .unwrap_or(0);
        if already > 0 { return; }

        let da = da_ref.borrow();
        // Get the SQL query — designer stores as "selectcommand", constructor as "selectcommandtext"
        let sql = da.fields.get("selectcommandtext")
            .or_else(|| da.fields.get("selectcommand"))
            .map(|v| v.as_string())
            .unwrap_or_default();
        let conn_str = da.fields.get("connectionstring")
            .map(|v| v.as_string())
            .unwrap_or_default();
        let conn_id_existing = da.fields.get("__conn_id")
            .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
            .unwrap_or(0);
        // Also check DataMember on the parent BindingSource if no SQL is set
        drop(da);

        if conn_str.is_empty() && conn_id_existing == 0 {
            eprintln!("[auto_fill] DataAdapter has no ConnectionString and no connection");
            return;
        }

        let dam = crate::data_access::get_global_dam();
        let mut dam_lock = dam.lock().unwrap();

        // Open connection if we don't have one yet
        let conn_id = if conn_id_existing > 0 {
            conn_id_existing
        } else if !conn_str.is_empty() {
            match dam_lock.open_connection(&conn_str) {
                Ok(id) => {
                    da_ref.borrow_mut().fields.insert("__conn_id".to_string(), Value::Long(id as i64));
                    id
                }
                Err(e) => {
                    eprintln!("[auto_fill] Failed to connect: {}", e);
                    return;
                }
            }
        } else {
            return;
        };

        if sql.is_empty() {
            eprintln!("[auto_fill] DataAdapter has no SelectCommand");
            return;
        }

        match dam_lock.execute_reader(conn_id, &sql) {
            Ok(rs_id) => {
                eprintln!("[auto_fill] OK: rs_id={}, rows={}",
                    rs_id,
                    dam_lock.recordsets.get(&rs_id).map(|r| r.record_count()).unwrap_or(0));
                da_ref.borrow_mut().fields.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
            }
            Err(e) => {
                eprintln!("[auto_fill] Query error: {}", e);
            }
        }
    }

    /// Get the row count from a DataSource (DataTable, DataSet, DataAdapter, or Array).
    /// If the DataSource is a DataAdapter that hasn't been filled yet, auto-fill it.
    pub fn binding_source_row_count(&self, datasource: &Value) -> i32 {
        let datasource = match datasource {
            Value::String(proxy_name) => &self.resolve_control_as_sender(proxy_name),
            _ => datasource,
        };
        match datasource {
            Value::Object(obj_ref) => {
                let obj = obj_ref.borrow();
                let dt_type = obj.fields.get("__type")
                    .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                    .unwrap_or_default();
                if dt_type == "DataTable" {
                    let rs_id = obj.fields.get("__rs_id")
                        .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                        .unwrap_or(0);
                    let dam = crate::data_access::get_global_dam();
                    let dam_lock = dam.lock().unwrap();
                    if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                        return rs.record_count();
                    }
                } else if dt_type == "DataSet" {
                    // Use first table if available
                    if let Some(Value::Array(tables)) = obj.fields.get("__tables") {
                        if let Some(Value::Object(dt_ref)) = tables.first() {
                            let dt = dt_ref.borrow();
                            let rs_id = dt.fields.get("__rs_id")
                                .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                .unwrap_or(0);
                            drop(dt);
                            drop(obj);
                            let dam = crate::data_access::get_global_dam();
                            let dam_lock = dam.lock().unwrap();
                            if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                                return rs.record_count();
                            }
                        }
                    }
                } else if dt_type == "DataAdapter" {
                    drop(obj);
                    Self::auto_fill_data_adapter(obj_ref);
                    let rs_id = obj_ref.borrow().fields.get("__rs_id")
                        .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                        .unwrap_or(0);
                    if rs_id > 0 {
                        let dam = crate::data_access::get_global_dam();
                        let dam_lock = dam.lock().unwrap();
                        if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                            return rs.record_count();
                        }
                    }
                } else if dt_type == "BindingSource" {
                    let inner_ds = obj.fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                    let data_member = obj.fields.get("datamember")
                        .map(|v| v.as_string()).unwrap_or_default();
                    drop(obj);
                    Self::inject_select_from_data_member(&inner_ds, &data_member);
                    return self.binding_source_row_count(&inner_ds);
                }
                0
            }
            Value::Array(arr) => arr.len() as i32,
            _ => 0,
        }
    }

    /// Get a DataRow from a DataSource at a given position.
    fn binding_source_get_row(&self, datasource: &Value, position: i32) -> Value {
        let datasource = match datasource {
            Value::String(proxy_name) => &self.resolve_control_as_sender(proxy_name),
            _ => datasource,
        };
        match datasource {
            Value::Object(obj_ref) => {
                let dt_type = {
                    let obj = obj_ref.borrow();
                    obj.fields.get("__type")
                        .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                        .unwrap_or_default()
                };

                // BindingSource: follow __datasource recursively
                if dt_type == "BindingSource" {
                    let obj = obj_ref.borrow();
                    let inner_ds = obj.fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                    let data_member = obj.fields.get("datamember")
                        .map(|v| v.as_string()).unwrap_or_default();
                    drop(obj);
                    Self::inject_select_from_data_member(&inner_ds, &data_member);
                    return self.binding_source_get_row(&inner_ds, position);
                }

                // DataAdapter: auto-fill if needed
                if dt_type == "DataAdapter" {
                    Self::auto_fill_data_adapter(obj_ref);
                }

                let rs_id = {
                    let obj = obj_ref.borrow();
                    if dt_type == "DataTable" || dt_type == "DataAdapter" {
                        obj.fields.get("__rs_id")
                            .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                            .unwrap_or(0)
                    } else if dt_type == "DataSet" {
                        if let Some(Value::Array(tables)) = obj.fields.get("__tables") {
                            if let Some(Value::Object(dt_ref)) = tables.first() {
                                dt_ref.borrow().fields.get("__rs_id")
                                    .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                    .unwrap_or(0)
                            } else { 0 }
                        } else { 0 }
                    } else { 0 }
                };

                if rs_id > 0 {
                    let dam = crate::data_access::get_global_dam();
                    let dam_lock = dam.lock().unwrap();
                    if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                        if let Some(db_row) = rs.rows.get(position as usize) {
                            let mut flds = std::collections::HashMap::new();
                            flds.insert("__type".to_string(), Value::String("DataRow".to_string()));
                            flds.insert("__rs_id".to_string(), Value::Long(rs_id as i64));
                            flds.insert("__row_index".to_string(), Value::Integer(position));
                            for (ci, col) in db_row.columns.iter().enumerate() {
                                let v = db_row.values.get(ci).cloned().unwrap_or_default();
                                flds.insert(col.to_lowercase(), Value::String(v));
                            }
                            let obj = crate::value::ObjectData { class_name: "DataRow".to_string(), fields: flds };
                            return Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj)));
                        }
                    }
                }
                Value::Nothing
            }
            Value::Array(arr) => {
                if let Some(val) = arr.get(position as usize) {
                    val.clone()
                } else {
                    Value::Nothing
                }
            }
            _ => Value::Nothing,
        }
    }

    /// Get all columns and rows from a DataSource for rendering in DataGridView.
    fn get_datasource_table_data(&self, datasource: &Value) -> (Vec<String>, Vec<Vec<String>>) {
        match datasource {
            Value::Object(obj_ref) => {
                let dt_type = {
                    let obj = obj_ref.borrow();
                    obj.fields.get("__type")
                        .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                        .unwrap_or_default()
                };

                // BindingSource: follow __datasource recursively
                if dt_type == "BindingSource" {
                    let obj = obj_ref.borrow();
                    let inner_ds = obj.fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                    let data_member = obj.fields.get("datamember")
                        .map(|v| v.as_string()).unwrap_or_default();
                    drop(obj);
                    Self::inject_select_from_data_member(&inner_ds, &data_member);
                    return self.get_datasource_table_data(&inner_ds);
                }

                // DataAdapter: auto-fill if needed
                if dt_type == "DataAdapter" {
                    Self::auto_fill_data_adapter(obj_ref);
                }

                let rs_id = {
                    let obj = obj_ref.borrow();
                    if dt_type == "DataTable" || dt_type == "DataAdapter" {
                        obj.fields.get("__rs_id")
                            .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                            .unwrap_or(0)
                    } else if dt_type == "DataSet" {
                        if let Some(Value::Array(tables)) = obj.fields.get("__tables") {
                            if let Some(Value::Object(dt_ref)) = tables.first() {
                                dt_ref.borrow().fields.get("__rs_id")
                                    .and_then(|v| if let Value::Long(l) = v { Some(*l as u64) } else { None })
                                    .unwrap_or(0)
                            } else { 0 }
                        } else { 0 }
                    } else { 0 }
                };

                if rs_id > 0 {
                    let dam = crate::data_access::get_global_dam();
                    let dam_lock = dam.lock().unwrap();
                    if let Some(rs) = dam_lock.recordsets.get(&rs_id) {
                        let columns = rs.columns.clone();
                        let rows: Vec<Vec<String>> = rs.rows.iter()
                            .map(|r| r.values.clone())
                            .collect();
                        return (columns, rows);
                    }
                }
                (Vec::new(), Vec::new())
            }
            Value::Array(arr) => {
                // Array of objects: infer columns from first element
                if let Some(Value::Object(first)) = arr.first() {
                    let first_borrow = first.borrow();
                    let mut columns: Vec<String> = first_borrow.fields.keys()
                        .filter(|k| !k.starts_with("__"))
                        .cloned()
                        .collect();
                    columns.sort();
                    let rows: Vec<Vec<String>> = arr.iter().map(|item| {
                        if let Value::Object(obj_r) = item {
                            let b = obj_r.borrow();
                            columns.iter().map(|col| {
                                b.fields.get(col).map(|v| v.as_string()).unwrap_or_default()
                            }).collect()
                        } else {
                            vec![item.as_string()]
                        }
                    }).collect();
                    (columns, rows)
                } else if let Some(Value::String(_)) = arr.first() {
                    let rows: Vec<Vec<String>> = arr.iter().map(|v| vec![v.as_string()]).collect();
                    (vec!["Value".to_string()], rows)
                } else {
                    (Vec::new(), Vec::new())
                }
            }
            _ => (Vec::new(), Vec::new()),
        }
    }

    fn expr_to_string(&self, expr: &Expression) -> String {
        match expr {
            Expression::Variable(name) => name.as_str().to_string(),
            Expression::Me => "Me".to_string(),
            Expression::MyBase => "MyBase".to_string(),
            Expression::Cast { expr, .. } => self.expr_to_string(expr),
            Expression::MemberAccess(obj, member) => {
                format!("{}.{}", self.expr_to_string(obj), member.as_str())
            }
            _ => "[expr]".to_string(),
        }
    }
    fn dispatch_file_method(&mut self, method_name: &str, args: &[Value]) -> Result<Value, RuntimeError> {
        match method_name.to_lowercase().as_str() {
            "exists" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let exists = std::path::Path::new(&path).exists();
                 Ok(Value::Boolean(exists))
            }
            "readalltext" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let content = std::fs::read_to_string(&path).map_err(|e| RuntimeError::Custom(format!("Error reading file: {}", e)))?;
                 Ok(Value::String(content))
            }
            "writealltext" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let content = args.get(1).ok_or(RuntimeError::Custom("Missing content argument".to_string()))?.as_string();
                 std::fs::write(&path, content).map_err(|e| RuntimeError::Custom(format!("Error writing file: {}", e)))?;
                 Ok(Value::Nothing)
            }
            "appendalltext" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let content = args.get(1).ok_or(RuntimeError::Custom("Missing content argument".to_string()))?.as_string();
                 let mut file = std::fs::OpenOptions::new().write(true).append(true).create(true).open(&path).map_err(|e| RuntimeError::Custom(format!("Error opening file: {}", e)))?;
                 use std::io::Write;
                 write!(file, "{}", content).map_err(|e| RuntimeError::Custom(format!("Error writing file: {}", e)))?;
                 Ok(Value::Nothing)
            }
            "copy" => {
                 let src = args.get(0).ok_or(RuntimeError::Custom("Missing source argument".to_string()))?.as_string();
                 let dest = args.get(1).ok_or(RuntimeError::Custom("Missing destination argument".to_string()))?.as_string();
                 let overwrite = args.get(2).map(|v| v.is_truthy()).unwrap_or(false);
                 if !overwrite && std::path::Path::new(&dest).exists() {
                     return Err(RuntimeError::Custom("Destination file exists".to_string()));
                 }
                 std::fs::copy(&src, &dest).map_err(|e| RuntimeError::Custom(format!("Error copying file: {}", e)))?;
                 Ok(Value::Nothing)
            }
            "delete" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 std::fs::remove_file(&path).map_err(|e| RuntimeError::Custom(format!("Error deleting file: {}", e)))?;
                 Ok(Value::Nothing)
            }
            "move" => {
                 let src = args.get(0).ok_or(RuntimeError::Custom("Missing source argument".to_string()))?.as_string();
                 let dest = args.get(1).ok_or(RuntimeError::Custom("Missing destination argument".to_string()))?.as_string();
                 std::fs::rename(&src, &dest).map_err(|e| RuntimeError::Custom(format!("Error moving file: {}", e)))?;
                 Ok(Value::Nothing)
            }
            "readalllines" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let content = std::fs::read_to_string(&path).map_err(|e| RuntimeError::Custom(format!("Error reading file: {}", e)))?;
                 let lines: Vec<Value> = content.lines().map(|l| Value::String(l.to_string())).collect();
                 Ok(Value::Array(lines))
            }
            "writealllines" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let lines = match args.get(1) {
                     Some(Value::Array(arr)) => arr.iter().map(|v| v.as_string()).collect::<Vec<_>>().join("\n"),
                     _ => return Err(RuntimeError::Custom("WriteAllLines requires an array argument".to_string())),
                 };
                 std::fs::write(&path, lines).map_err(|e| RuntimeError::Custom(format!("Error writing file: {}", e)))?;
                 Ok(Value::Nothing)
            }
            "readallbytes" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let bytes = std::fs::read(&path).map_err(|e| RuntimeError::Custom(format!("Error reading file: {}", e)))?;
                 let arr: Vec<Value> = bytes.into_iter().map(|b| Value::Byte(b)).collect();
                 Ok(Value::Array(arr))
            }
            "writeallbytes" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let bytes: Vec<u8> = match args.get(1) {
                     Some(Value::Array(arr)) => arr.iter().map(|v| match v {
                         Value::Byte(b) => *b,
                         Value::Integer(i) => *i as u8,
                         _ => 0u8,
                     }).collect(),
                     _ => return Err(RuntimeError::Custom("WriteAllBytes requires a byte array argument".to_string())),
                 };
                 std::fs::write(&path, &bytes).map_err(|e| RuntimeError::Custom(format!("Error writing file: {}", e)))?;
                 Ok(Value::Nothing)
            }
            "getlastwritetime" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let meta = std::fs::metadata(&path).map_err(|e| RuntimeError::Custom(format!("Error accessing file: {}", e)))?;
                 let modified = meta.modified().unwrap_or(std::time::SystemTime::UNIX_EPOCH);
                 let secs = modified.duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_secs() as i64;
                 let ndt = chrono::DateTime::from_timestamp(secs, 0).map(|dt| dt.naive_utc()).unwrap_or_default();
                 Ok(Value::Date(date_to_ole(ndt)))
            }
            "getcreationtime" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let meta = std::fs::metadata(&path).map_err(|e| RuntimeError::Custom(format!("Error accessing file: {}", e)))?;
                 let created = meta.created().unwrap_or(std::time::SystemTime::UNIX_EPOCH);
                 let secs = created.duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_secs() as i64;
                 let ndt = chrono::DateTime::from_timestamp(secs, 0).map(|dt| dt.naive_utc()).unwrap_or_default();
                 Ok(Value::Date(date_to_ole(ndt)))
            }
            _ => Err(RuntimeError::UndefinedFunction(format!("System.IO.File.{}", method_name)))
        }
    }

    fn dispatch_directory_method(&mut self, method_name: &str, args: &[Value]) -> Result<Value, RuntimeError> {
        match method_name.to_lowercase().as_str() {
            "exists" => {
                let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                Ok(Value::Boolean(std::path::Path::new(&path).is_dir()))
            }
            "createdirectory" => {
                let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                std::fs::create_dir_all(&path).map_err(|e| RuntimeError::Custom(format!("Directory.CreateDirectory: {}", e)))?;
                Ok(Value::Nothing)
            }
            "delete" => {
                let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                let recursive = args.get(1).map(|v| v.is_truthy()).unwrap_or(false);
                if recursive {
                    std::fs::remove_dir_all(&path).map_err(|e| RuntimeError::Custom(format!("Directory.Delete: {}", e)))?;
                } else {
                    std::fs::remove_dir(&path).map_err(|e| RuntimeError::Custom(format!("Directory.Delete: {}", e)))?;
                }
                Ok(Value::Nothing)
            }
            "getfiles" => {
                let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                let pattern = args.get(1).map(|v| v.as_string()).unwrap_or_default();
                match std::fs::read_dir(&path) {
                    Ok(entries) => {
                        let files: Vec<Value> = entries
                            .filter_map(|e| e.ok())
                            .filter(|e| e.path().is_file())
                            .filter(|e| {
                                if pattern.is_empty() { return true; }
                                let name = e.file_name().to_string_lossy().to_lowercase();
                                let pat = pattern.replace("*.", ".").to_lowercase();
                                name.ends_with(&pat)
                            })
                            .map(|e| Value::String(e.path().to_string_lossy().to_string()))
                            .collect();
                        Ok(Value::Array(files))
                    }
                    Err(e) => Err(RuntimeError::Custom(format!("Directory.GetFiles: {}", e))),
                }
            }
            "getdirectories" => {
                let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                match std::fs::read_dir(&path) {
                    Ok(entries) => {
                        let dirs: Vec<Value> = entries
                            .filter_map(|e| e.ok())
                            .filter(|e| e.path().is_dir())
                            .map(|e| Value::String(e.path().to_string_lossy().to_string()))
                            .collect();
                        Ok(Value::Array(dirs))
                    }
                    Err(e) => Err(RuntimeError::Custom(format!("Directory.GetDirectories: {}", e))),
                }
            }
            "getcurrentdirectory" => {
                let cwd = std::env::current_dir().unwrap_or_default().to_string_lossy().to_string();
                Ok(Value::String(cwd))
            }
            "setcurrentdirectory" => {
                let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                std::env::set_current_dir(&path).map_err(|e| RuntimeError::Custom(format!("Directory.SetCurrentDirectory: {}", e)))?;
                Ok(Value::Nothing)
            }
            "move" => {
                let src = args.get(0).ok_or(RuntimeError::Custom("Missing source argument".to_string()))?.as_string();
                let dest = args.get(1).ok_or(RuntimeError::Custom("Missing destination argument".to_string()))?.as_string();
                std::fs::rename(&src, &dest).map_err(|e| RuntimeError::Custom(format!("Directory.Move: {}", e)))?;
                Ok(Value::Nothing)
            }
            "getparent" => {
                let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                let parent = std::path::Path::new(&path).parent().unwrap_or(std::path::Path::new("")).to_string_lossy().to_string();
                Ok(Value::String(parent))
            }
            "getdirectoryroot" => {
                let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                // On Unix, root is always "/"; on Windows it's "C:\"
                let root = if cfg!(target_os = "windows") {
                    path.get(..3).unwrap_or("C:\\").to_string()
                } else {
                    "/".to_string()
                };
                Ok(Value::String(root))
            }
            _ => Err(RuntimeError::UndefinedFunction(format!("System.IO.Directory.{}", method_name)))
        }
    }

    fn dispatch_path_method(&mut self, method_name: &str, args: &[Value]) -> Result<Value, RuntimeError> {
        match method_name.to_lowercase().as_str() {
            "combine" => {
                 let mut path = std::path::PathBuf::new();
                 for arg in args {
                     path.push(arg.as_string());
                 }
                 Ok(Value::String(path.to_string_lossy().to_string()))
            }
            "getfilename" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 Ok(Value::String(std::path::Path::new(&path).file_name().unwrap_or_default().to_string_lossy().to_string()))
            }
            "getextension" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let ext = std::path::Path::new(&path).extension().unwrap_or_default().to_string_lossy().to_string();
                 if ext.is_empty() { Ok(Value::String(String::new())) }
                 else { Ok(Value::String(format!(".{}", ext))) }
            }
            "getdirectoryname" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 Ok(Value::String(std::path::Path::new(&path).parent().unwrap_or(std::path::Path::new("")).to_string_lossy().to_string()))
            }
            "getfilenamewithoutextension" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 Ok(Value::String(std::path::Path::new(&path).file_stem().unwrap_or_default().to_string_lossy().to_string()))
            }
            "getfullpath" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let full = std::fs::canonicalize(&path)
                     .unwrap_or_else(|_| std::path::PathBuf::from(&path));
                 Ok(Value::String(full.to_string_lossy().to_string()))
            }
            "gettemppath" => {
                 Ok(Value::String(std::env::temp_dir().to_string_lossy().to_string()))
            }
            "gettempfilename" => {
                 let tmp = std::env::temp_dir();
                 let name = format!("tmp{:x}.tmp", std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_nanos());
                 let path = tmp.join(name);
                 // Create the file like .NET does
                 std::fs::write(&path, "").ok();
                 Ok(Value::String(path.to_string_lossy().to_string()))
            }
            "hasextension" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 Ok(Value::Boolean(std::path::Path::new(&path).extension().is_some()))
            }
            "ispathrooted" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 Ok(Value::Boolean(std::path::Path::new(&path).is_absolute()))
            }
            "changeextension" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let ext = args.get(1).ok_or(RuntimeError::Custom("Missing extension argument".to_string()))?.as_string();
                 let mut p = std::path::PathBuf::from(&path);
                 p.set_extension(ext.trim_start_matches('.'));
                 Ok(Value::String(p.to_string_lossy().to_string()))
            }
            "directoryseparatorchar" => {
                 Ok(Value::String(std::path::MAIN_SEPARATOR.to_string()))
            }
            _ => Err(RuntimeError::UndefinedFunction(format!("System.IO.Path.{}", method_name)))
        }
    }

    fn dispatch_console_method(&mut self, method_name: &str, args: &[Value]) -> Result<Value, RuntimeError> {
        match method_name.to_lowercase().as_str() {
            "write" | "writeline" => {
                let msg = crate::builtins::console_fns::console_write_fn(args);
                let final_msg = if method_name.eq_ignore_ascii_case("writeline") {
                    format!("{}\n", msg)
                } else {
                    msg
                };
                self.send_console_output(final_msg);
                Ok(Value::Nothing)
            }
            "readline" | "readkey" => {
                // Interactive Console.ReadLine: send request, block for response
                if let Some(tx) = &self.console_tx {
                    let _ = tx.send(crate::ConsoleMessage::InputRequest);
                    if let Some(rx) = &self.console_input_rx {
                        match rx.recv() {
                            Ok(input) => Ok(Value::String(input)),
                            Err(_) => Ok(Value::String(String::new())),
                        }
                    } else {
                        Ok(Value::String(String::new()))
                    }
                } else {
                    // Check if stdin is a TTY (CLI with real terminal)
                    use std::io::IsTerminal;
                    if std::io::stdin().is_terminal() {
                        // Real terminal: read from stdin
                        let mut line = String::new();
                        match std::io::stdin().read_line(&mut line) {
                            Ok(_) => {
                                if line.ends_with('\n') { line.pop(); }
                                if line.ends_with('\r') { line.pop(); }
                                Ok(Value::String(line))
                            }
                            Err(_) => Ok(Value::String(String::new())),
                        }
                    } else if std::io::stdin().lock().fill_buf().map(|b| !b.is_empty()).unwrap_or(false) {
                        // Piped stdin with data available
                        let mut line = String::new();
                        match std::io::stdin().read_line(&mut line) {
                            Ok(_) => {
                                if line.ends_with('\n') { line.pop(); }
                                if line.ends_with('\r') { line.pop(); }
                                Ok(Value::String(line))
                            }
                            Err(_) => Ok(Value::String(String::new())),
                        }
                    } else {
                        // GUI mode: no console channel, no real terminal — use native input dialog 
                        match crate::builtins::info_fns::show_native_input_dialog("Console.ReadLine", "Input", "") {
                            Some(input) => Ok(Value::String(input)),
                            None => Ok(Value::String(String::new())),
                        }
                    }
                }
            }
            "read" => {
                // Console.Read() — returns a single character as an integer
                if let Some(tx) = &self.console_tx {
                    let _ = tx.send(crate::ConsoleMessage::InputRequest);
                    if let Some(rx) = &self.console_input_rx {
                        match rx.recv() {
                            Ok(input) => {
                                if let Some(ch) = input.chars().next() {
                                    Ok(Value::Integer(ch as i32))
                                } else {
                                    Ok(Value::Integer(-1))
                                }
                            }
                            Err(_) => Ok(Value::Integer(-1)),
                        }
                    } else {
                        Ok(Value::Integer(-1))
                    }
                } else {
                    let mut buf = [0u8; 1];
                    match std::io::Read::read(&mut std::io::stdin(), &mut buf) {
                        Ok(1) => Ok(Value::Integer(buf[0] as i32)),
                        _ => Ok(Value::Integer(-1)),
                    }
                }
            }
            "clear" => {
                if let Some(tx) = &self.console_tx {
                    let _ = tx.send(crate::ConsoleMessage::Clear);
                } else {
                    self.side_effects.push_back(crate::RuntimeSideEffect::ConsoleClear);
                }
                Ok(Value::Nothing)
            }
            "resetcolor" => {
                // Reset Console colors to defaults (Gray on Black)
                if let Ok(Value::Object(ref obj)) = self.env.get("console") {
                    let mut fields = obj.borrow_mut();
                    fields.fields.insert("foregroundcolor".to_string(), Value::Integer(7));
                    fields.fields.insert("backgroundcolor".to_string(), Value::Integer(0));
                }
                Ok(Value::Nothing)
            }
            "setcursorposition" => {
                // Console.SetCursorPosition(left, top)
                let left = args.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                let top = args.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                if let Ok(Value::Object(ref obj)) = self.env.get("console") {
                    let mut fields = obj.borrow_mut();
                    fields.fields.insert("cursorleft".to_string(), Value::Integer(left));
                    fields.fields.insert("cursortop".to_string(), Value::Integer(top));
                }
                Ok(Value::Nothing)
            }
            "beep" => {
                // Console.Beep() — best-effort; no-op in UI, print BEL in CLI
                if self.console_tx.is_none() {
                    print!("\x07");
                }
                Ok(Value::Nothing)
            }
            "opensandardinput" | "openstandardoutput" | "openstandarderror" => {
                // Stub — return Nothing for stream methods
                Ok(Value::Nothing)
            }
            _ => Err(RuntimeError::UndefinedFunction(format!("System.Console.{}", method_name)))
        }
    }

    fn dispatch_application_method(&mut self, method_name: &str, args: &[Value]) -> Result<Value, RuntimeError> {
        match method_name.to_lowercase().as_str() {
            "run" => {
                if let Some(arg) = args.first() {
                    if let Value::Object(obj_ref) = arg {
                        let class_name = obj_ref.borrow().class_name.clone();
                        // Set as the active form instance
                        self.env.define_global("__form_instance__", arg.clone());
                        self.side_effects.push_back(crate::RuntimeSideEffect::RunApplication { form_name: class_name });
                    }
                }
                Ok(Value::Nothing)
            }
            "exit" => {
                // Close the current form if any
                if let Ok(Value::Object(obj_ref)) = self.env.get("__form_instance__") {
                    let borrow = obj_ref.borrow();
                    let name = borrow.fields.get("name")
                        .map(|v| v.as_string())
                        .filter(|s| !s.is_empty())
                        .unwrap_or(borrow.class_name.clone());
                        
                    if !name.is_empty() {
                         self.side_effects.push_back(crate::RuntimeSideEffect::FormClose { form_name: name });
                    }
                }
                Ok(Value::Nothing)
            }
            "doevents" => {
                // No-op in this interpreter model (events are handled by UI thread)
                Ok(Value::Nothing)
            }
            _ => Err(RuntimeError::UndefinedFunction(format!("Application.{}", method_name)))
        }
    }

    fn dispatch_math_method(&mut self, method_name: &str, args: &[Value]) -> Result<Value, RuntimeError> {
        match method_name.to_lowercase().as_str() {
            "abs" => crate::builtins::math_fns::abs_fn(args),
            "max" => crate::builtins::math_fns::max_fn(args),
            "min" => crate::builtins::math_fns::min_fn(args),
            "sqrt" | "sqr" => crate::builtins::math_fns::sqr_fn(args),
            "round" => crate::builtins::math_fns::round_fn(args),
            "floor" => crate::builtins::math_fns::floor_fn(args),
            "truncate" => {
                let val = args.get(0).ok_or(RuntimeError::Custom("Math.Truncate requires an argument".to_string()))?.as_double()?;
                Ok(Value::Double(val.trunc()))
            }
            "ceiling" => crate::builtins::math_fns::ceiling_fn(args),
            "pow" => crate::builtins::math_fns::pow_fn(args),
            "exp" => crate::builtins::math_fns::exp_fn(args),
            "log" => {
                if args.len() >= 2 {
                    // Math.Log(value, base)
                    let val = args[0].as_double()?;
                    let base = args[1].as_double()?;
                    Ok(Value::Double(val.log(base)))
                } else {
                    crate::builtins::math_fns::log_fn(args)
                }
            }
            "log10" => {
                let val = args.get(0).ok_or(RuntimeError::Custom("Math.Log10 requires an argument".to_string()))?.as_double()?;
                Ok(Value::Double(val.log10()))
            }
            "log2" => {
                let val = args.get(0).ok_or(RuntimeError::Custom("Math.Log2 requires an argument".to_string()))?.as_double()?;
                Ok(Value::Double(val.log2()))
            }
            "sin" => crate::builtins::math_fns::sin_fn(args),
            "cos" => crate::builtins::math_fns::cos_fn(args),
            "tan" => crate::builtins::math_fns::tan_fn(args),
            "asin" => {
                let val = args.get(0).ok_or(RuntimeError::Custom("Math.Asin requires an argument".to_string()))?.as_double()?;
                Ok(Value::Double(val.asin()))
            }
            "acos" => {
                let val = args.get(0).ok_or(RuntimeError::Custom("Math.Acos requires an argument".to_string()))?.as_double()?;
                Ok(Value::Double(val.acos()))
            }
            "atan" | "atn" => crate::builtins::math_fns::atn_fn(args),
            "atan2" => crate::builtins::math_fns::atan2_fn(args),
            "sign" | "sgn" => crate::builtins::math_fns::sgn_fn(args),
            "sinh" => {
                let val = args.get(0).ok_or(RuntimeError::Custom("Math.Sinh requires an argument".to_string()))?.as_double()?;
                Ok(Value::Double(val.sinh()))
            }
            "cosh" => {
                let val = args.get(0).ok_or(RuntimeError::Custom("Math.Cosh requires an argument".to_string()))?.as_double()?;
                Ok(Value::Double(val.cosh()))
            }
            "tanh" => {
                let val = args.get(0).ok_or(RuntimeError::Custom("Math.Tanh requires an argument".to_string()))?.as_double()?;
                Ok(Value::Double(val.tanh()))
            }
            "clamp" => {
                let val = args.get(0).ok_or(RuntimeError::Custom("Math.Clamp requires 3 arguments".to_string()))?.as_double()?;
                let min = args.get(1).ok_or(RuntimeError::Custom("Math.Clamp requires 3 arguments".to_string()))?.as_double()?;
                let max = args.get(2).ok_or(RuntimeError::Custom("Math.Clamp requires 3 arguments".to_string()))?.as_double()?;
                Ok(Value::Double(val.clamp(min, max)))
            }
            "pi" => Ok(Value::Double(std::f64::consts::PI)),
            "e" => Ok(Value::Double(std::f64::consts::E)),
            _ => Err(RuntimeError::UndefinedFunction(format!("System.Math.{}", method_name)))
        }

    }

    fn call_lambda(&mut self, lambda_val: Value, args: &[Value]) -> Result<Value, RuntimeError> {
        if let Value::Lambda { params, body, env } = lambda_val {
            if args.len() != params.len() {
                return Err(RuntimeError::Custom(format!("Lambda expects {} arguments, got {}", params.len(), args.len())));
            }
            
            // Save current environment
            let mut prev_env = self.env.clone();
            
            // Switch to captured environment (Snapshot)
            self.env = env.borrow().clone();
            self.env.push_scope();
            
            // Bind arguments
            let param_names: std::collections::HashSet<String> = params.iter()
                .map(|p| p.name.as_str().to_lowercase())
                .collect();
            for (param, arg) in params.iter().zip(args.iter()) {
                self.env.define(param.name.as_str(), arg.clone());
            }
            
            let result = match &*body {
                vybe_parser::ast::expr::LambdaBody::Expression(expr) => {
                    self.evaluate_expr(expr)
                }
                vybe_parser::ast::expr::LambdaBody::Statement(stmt) => {
                    match self.execute(stmt) {
                        Ok(_) => Ok(Value::Nothing),
                        Err(RuntimeError::Return(val)) => Ok(val.unwrap_or(Value::Nothing)),
                        Err(e) => Err(e),
                    }
                }
                vybe_parser::ast::expr::LambdaBody::Block(stmts) => {
                    let mut final_res = Ok(Value::Nothing);
                    for stmt in stmts {
                        match self.execute(stmt) {
                            Ok(_) => {},
                            Err(RuntimeError::Return(val)) => {
                                final_res = Ok(val.unwrap_or(Value::Nothing));
                                break;
                            }
                            Err(e) => {
                                final_res = Err(e);
                                break;
                            }
                        }
                    }
                    final_res
                }
            };
            
            // Propagate modified closure variables back (VB.NET captures by reference)
            // Skip lambda parameters — only propagate outer variables that the lambda modified
            for (name, value) in self.env.all_variables() {
                if param_names.contains(&name) { continue; }
                // Update captured env (so future calls see changes)
                let _ = env.borrow_mut().set(&name, value.clone());
                // Update caller env (so caller sees changes)
                if prev_env.get(&name).is_ok() {
                    let _ = prev_env.set(&name, value);
                }
            }
            
            // Restore environment
            self.env = prev_env;
            result
        } else {
            Err(RuntimeError::Custom("Not a lambda".to_string()))
        }
    }
}

impl Default for Interpreter {
    fn default() -> Self {
        Self::new()
    }
}

/// Compare two Values for ordering. Numbers compare numerically, strings lexically.
fn compare_values_ordering(a: &Value, b: &Value) -> std::cmp::Ordering {
    // Try to extract numeric values for comparison
    let a_num = match a {
        Value::Integer(n) => Some(*n as f64),
        Value::Long(n) => Some(*n as f64),
        Value::Single(n) => Some(*n as f64),
        Value::Double(n) => Some(*n),
        _ => None,
    };
    let b_num = match b {
        Value::Integer(n) => Some(*n as f64),
        Value::Long(n) => Some(*n as f64),
        Value::Single(n) => Some(*n as f64),
        Value::Double(n) => Some(*n),
        _ => None,
    };
    match (a_num, b_num) {
        (Some(x), Some(y)) => x.partial_cmp(&y).unwrap_or(std::cmp::Ordering::Equal),
        _ => a.as_string().cmp(&b.as_string()),
    }
}

/// Create a multi-dimensional array as nested Value::Array.
/// dims = [3, 4] creates a 3-element array where each element is a 4-element array.
fn create_multi_dim_array(dims: &[usize], default: &Value) -> Value {
    if dims.len() == 1 {
        Value::Array(vec![default.clone(); dims[0]])
    } else {
        let inner = create_multi_dim_array(&dims[1..], default);
        Value::Array(vec![inner; dims[0]])
    }
}

/// Set an element in a multi-dimensional (nested) array.
fn set_multi_dim_element(arr: &mut Value, indices: &[usize], val: Value) -> Result<(), RuntimeError> {
    if indices.len() == 1 {
        if let Value::Array(vec) = arr {
            if indices[0] < vec.len() {
                vec[indices[0]] = val;
                return Ok(());
            }
            return Err(RuntimeError::Custom("Array index out of bounds".to_string()));
        }
        return Err(RuntimeError::Custom("Not an array".to_string()));
    }
    // Recurse into the inner dimension
    if let Value::Array(vec) = arr {
        if indices[0] < vec.len() {
            return set_multi_dim_element(&mut vec[indices[0]], &indices[1..], val);
        }
        return Err(RuntimeError::Custom("Array index out of bounds".to_string()));
    }
    Err(RuntimeError::Custom("Not an array".to_string()))
}

fn default_value_for_type(_name: &str, var_type: &Option<vybe_parser::VBType>) -> Value {
    match var_type {
        Some(vybe_parser::VBType::Integer) => Value::Integer(0),
        Some(vybe_parser::VBType::Long) => Value::Long(0),
        Some(vybe_parser::VBType::Single) => Value::Single(0.0),
        Some(vybe_parser::VBType::Double) => Value::Double(0.0),
        Some(vybe_parser::VBType::String) => Value::String(String::new()),
        Some(vybe_parser::VBType::Boolean) => Value::Boolean(false),
        Some(vybe_parser::VBType::Variant) => Value::Nothing,
        _ => Value::Nothing,
    }
}

// ===== OLE Automation Date helpers =====
// OLE Automation Date: f64 representing days since 1899-12-30
// Integer part = days, fractional part = time of day

fn date_to_ole(dt: chrono::NaiveDateTime) -> f64 {
    let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0, 0, 0).unwrap();
    let dur = dt.signed_duration_since(base);
    let days = dur.num_days() as f64;
    let secs = (dur.num_seconds() % 86400) as f64;
    days + (secs / 86400.0)
}

fn ole_to_dt(ole: f64) -> chrono::NaiveDateTime {
    let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0, 0, 0).unwrap();
    let days = ole.trunc() as i64;
    let frac = ole.fract();
    let secs = (frac * 86400.0).round() as i64;
    base.checked_add_signed(chrono::Duration::days(days))
        .and_then(|d| d.checked_add_signed(chrono::Duration::seconds(secs)))
        .unwrap_or(base)
}

fn ymd_to_ole(y: i32, m: u32, d: u32, h: u32, min: u32, s: u32) -> f64 {
    let dt = chrono::NaiveDate::from_ymd_opt(y, m, d)
        .and_then(|date| date.and_hms_opt(h, min, s))
        .unwrap_or_else(|| chrono::NaiveDate::from_ymd_opt(2000, 1, 1).unwrap().and_hms_opt(0, 0, 0).unwrap());
    date_to_ole(dt)
}

fn now_ole() -> f64 {
    date_to_ole(chrono::Local::now().naive_local())
}

fn today_ole() -> f64 {
    let now = chrono::Local::now().naive_local();
    let today = now.date().and_hms_opt(0, 0, 0).unwrap();
    date_to_ole(today)
}

fn utcnow_ole() -> f64 {
    date_to_ole(chrono::Utc::now().naive_utc())
}

fn parse_date_to_ole(s: &str) -> Option<f64> {
    let s = s.trim();
    let formats = [
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M", "%Y-%m-%d %H:%M", "%m/%d/%Y", "%Y-%m-%d",
        "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%dT%H:%M:%S%.f", "%m/%d/%Y %I:%M:%S %p",
    ];
    for fmt in &formats {
        if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(s, fmt) {
            return Some(date_to_ole(dt));
        }
    }
    for fmt in &["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"] {
        if let Ok(d) = chrono::NaiveDate::parse_from_str(s, fmt) {
            return Some(date_to_ole(d.and_hms_opt(0, 0, 0).unwrap()));
        }
    }
    None
}

fn format_ole_date(ole: f64, fmt: &str) -> String {
    let dt = ole_to_dt(ole);
    if fmt.is_empty() {
        return dt.format("%m/%d/%Y %H:%M:%S").to_string();
    }

    // Standard single-letter formats
    if fmt.len() == 1 {
        match fmt {
            "d" => return dt.format("%m/%d/%Y").to_string(), // Short date
            "D" => return dt.format("%A, %B %d, %Y").to_string(), // Long date
            "t" => return dt.format("%H:%M").to_string(), // Short time
            "T" => return dt.format("%H:%M:%S").to_string(), // Long time
            "f" => return dt.format("%A, %B %d, %Y %H:%M").to_string(), // Full date/time (short time)
            "F" => return dt.format("%A, %B %d, %Y %H:%M:%S").to_string(), // Full date/time (long time)
            "g" => return dt.format("%m/%d/%Y %H:%M").to_string(), // General date (short time)
            "G" => return dt.format("%m/%d/%Y %H:%M:%S").to_string(), // General date (long time)
            "M" | "m" => return dt.format("%B %d").to_string(), // Month/Day
            "Y" | "y" => return dt.format("%B, %Y").to_string(), // Year/Month
            "s" => return dt.format("%Y-%m-%dT%H:%M:%S").to_string(), // Sortable
            "u" => return dt.format("%Y-%m-%d %H:%M:%SZ").to_string(), // Universal sortable
            _ => {}
        }
    }

    // Convert .NET format strings to chrono format
    let chrono_fmt = fmt
        .replace("yyyy", "%Y").replace("yy", "%y")
        .replace("MMMM", "%B").replace("MMM", "%b").replace("MM", "%m")
        .replace("dddd", "%A").replace("ddd", "%a").replace("dd", "%d")
        .replace("HH", "%H").replace("hh", "%I")
        .replace("mm", "%M").replace("ss", "%S")
        .replace("tt", "%p").replace("fff", "%3f");
    dt.format(&chrono_fmt).to_string()
}

// Helper: get a folder path with env var fallback
fn dirs_fallback(env_var: &str, subdir: &str) -> String {
    std::env::var(env_var).unwrap_or_else(|_| {
        let home = std::env::var("HOME").unwrap_or_else(|_| "/tmp".to_string());
        format!("{}/{}", home, subdir)
    })
}

// Base64 encode without external crate
fn base64_encode(data: &[u8]) -> String {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut result = String::new();
    for chunk in data.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = if chunk.len() > 1 { chunk[1] as u32 } else { 0 };
        let b2 = if chunk.len() > 2 { chunk[2] as u32 } else { 0 };
        let triple = (b0 << 16) | (b1 << 8) | b2;
        result.push(CHARS[((triple >> 18) & 0x3F) as usize] as char);
        result.push(CHARS[((triple >> 12) & 0x3F) as usize] as char);
        if chunk.len() > 1 {
            result.push(CHARS[((triple >> 6) & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }
        if chunk.len() > 2 {
            result.push(CHARS[(triple & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }
    }
    result
}

// Base64 decode without external crate
fn base64_decode(input: &str) -> Result<Vec<u8>, String> {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut result = Vec::new();
    let s: Vec<u8> = input.bytes().filter(|b| *b != b'\n' && *b != b'\r' && *b != b' ').collect();
    if s.len() % 4 != 0 { return Err("Invalid base64 length".to_string()); }
    for chunk in s.chunks(4) {
        let vals: Vec<u32> = chunk.iter().map(|&c| {
            if c == b'=' { 0 }
            else { CHARS.iter().position(|&x| x == c).unwrap_or(0) as u32 }
        }).collect();
        let triple = (vals[0] << 18) | (vals[1] << 12) | (vals[2] << 6) | vals[3];
        result.push((triple >> 16) as u8);
        if chunk[2] != b'=' { result.push((triple >> 8) as u8); }
        if chunk[3] != b'=' { result.push(triple as u8); }
    }
    Ok(result)
}

/// VB.NET Like operator pattern matching
/// Supports: * (any chars), ? (single char), # (single digit), [charlist], [!charlist]
fn vb_like_match(text: &str, pattern: &str) -> bool {
    let text_chars: Vec<char> = text.chars().collect();
    let pattern_chars: Vec<char> = pattern.chars().collect();
    vb_like_match_inner(&text_chars, &pattern_chars)
}

fn vb_like_match_inner(text: &[char], pattern: &[char]) -> bool {
    if pattern.is_empty() {
        return text.is_empty();
    }
    match pattern[0] {
        '*' => {
            // Try matching rest of pattern at every position
            for i in 0..=text.len() {
                if vb_like_match_inner(&text[i..], &pattern[1..]) {
                    return true;
                }
            }
            false
        }
        '?' => {
            if text.is_empty() { return false; }
            vb_like_match_inner(&text[1..], &pattern[1..])
        }
        '#' => {
            if text.is_empty() || !text[0].is_ascii_digit() { return false; }
            vb_like_match_inner(&text[1..], &pattern[1..])
        }
        '[' => {
            if text.is_empty() { return false; }
            // Find closing ]
            if let Some(close) = pattern.iter().position(|&c| c == ']') {
                let inside = &pattern[1..close];
                let (negate, chars) = if !inside.is_empty() && inside[0] == '!' {
                    (true, &inside[1..])
                } else {
                    (false, inside)
                };
                let mut matches = false;
                let mut i = 0;
                while i < chars.len() {
                    if i + 2 < chars.len() && chars[i + 1] == '-' {
                        // Range: [a-z]
                        if text[0] >= chars[i] && text[0] <= chars[i + 2] {
                            matches = true;
                        }
                        i += 3;
                    } else {
                        if text[0] == chars[i] {
                            matches = true;
                        }
                        i += 1;
                    }
                }
                if negate { matches = !matches; }
                if matches {
                    vb_like_match_inner(&text[1..], &pattern[close + 1..])
                } else {
                    false
                }
            } else {
                // No closing bracket — treat [ as literal
                if text.is_empty() || text[0] != '[' { return false; }
                vb_like_match_inner(&text[1..], &pattern[1..])
            }
        }
        c => {
            if text.is_empty() { return false; }
            if text[0].to_ascii_lowercase() == c.to_ascii_lowercase() {
                vb_like_match_inner(&text[1..], &pattern[1..])
            } else {
                false
            }
        }
    }
}

/// Check whether a method body contains a `MyBase.New(...)` call.
/// Used to decide if automatic base-class constructor chaining is needed.
fn body_contains_mybase_new(body: &[vybe_parser::ast::Statement]) -> bool {
    use vybe_parser::ast::Expression;
    for stmt in body {
        if expr_in_stmt_matches(stmt, &|e| {
            matches!(e, Expression::MethodCall(obj, method, _)
                if matches!(obj.as_ref(), Expression::MyBase)
                   && method.as_str().eq_ignore_ascii_case("new"))
        }) {
            return true;
        }
    }
    false
}

/// Recursively check if any expression inside a statement satisfies a predicate.
fn expr_in_stmt_matches(stmt: &vybe_parser::ast::Statement, pred: &dyn Fn(&vybe_parser::ast::Expression) -> bool) -> bool {
    use vybe_parser::ast::Statement;
    match stmt {
        Statement::ExpressionStatement(e) => expr_matches(e, pred),
        Statement::Assignment { value, .. } => expr_matches(value, pred),
        Statement::MemberAssignment { object, value, .. } => expr_matches(object, pred) || expr_matches(value, pred),
        Statement::If { condition, then_branch, elseif_branches, else_branch, .. } => {
            if expr_matches(condition, pred) { return true; }
            for s in then_branch { if expr_in_stmt_matches(s, pred) { return true; } }
            for (c, b) in elseif_branches {
                if expr_matches(c, pred) { return true; }
                for s in b { if expr_in_stmt_matches(s, pred) { return true; } }
            }
            if let Some(eb) = else_branch {
                for s in eb { if expr_in_stmt_matches(s, pred) { return true; } }
            }
            false
        }
        _ => false, // For constructor chaining detection, other statement types are unlikely
    }
}

fn expr_matches(expr: &vybe_parser::ast::Expression, pred: &dyn Fn(&vybe_parser::ast::Expression) -> bool) -> bool {
    use vybe_parser::ast::Expression;
    if pred(expr) { return true; }
    match expr {
        Expression::MethodCall(obj, _, args) => {
            if expr_matches(obj, pred) { return true; }
            args.iter().any(|a| expr_matches(a, pred))
        }
        Expression::MemberAccess(obj, _) => expr_matches(obj, pred),
        Expression::Call(_, args) => args.iter().any(|a| expr_matches(a, pred)),
        _ => false,
    }
}
// ── Runtime Extensions ──

impl Interpreter {
    fn execute_query(&mut self, query: &vybe_parser::ast::query::QueryExpression) -> Result<Value, RuntimeError> {
        use vybe_parser::ast::query::*;
        
        let range = query.from_clause.ranges.first().ok_or(RuntimeError::Custom("Query must have range variable".to_string()))?;
        let collection_val = self.evaluate_expr(&range.collection)?;
        
        // Ensure we're working with a collection
        let items: Vec<Value> = match collection_val {
            Value::Array(a) => a.clone(),
            Value::Collection(c) => c.borrow().items.clone(),
            _ => return Err(RuntimeError::Custom("Query source must be array or collection".to_string()))
        };
        
        self.env.push_scope();
        
        let mut filtered_items = Vec::new();
        for item in items {
            self.env.define(&range.name, item.clone());
            let mut include = true;
            
            for clause in &query.body.clauses {
                match clause {
                    QueryClause::Where(expr) => {
                         if !self.evaluate_expr(expr)?.as_bool().unwrap_or(false) {
                             include = false;
                             break;
                         }
                    }
                    QueryClause::Let { name, value } => {
                        let v = self.evaluate_expr(value)?;
                        self.env.define(name, v);
                    }
                    _ => {} // OrderBy handled later
                }
            }
            
            if include {
                filtered_items.push(item); 
            }
        }
        
        // TODO: Sort (OrderBy)
        
        // Project (Select)
        let mut results = Vec::new();
        for item in filtered_items {
            self.env.define(&range.name, item.clone());
            // Re-run Let for context
            for clause in &query.body.clauses {
                if let QueryClause::Let { name, value } = clause {
                     let v = self.evaluate_expr(value)?;
                     self.env.define(name, v);
                }
            }
            
            match &query.body.select_or_group {
                SelectOrGroupClause::Select(exprs) => {
                    if exprs.len() == 1 {
                        results.push(self.evaluate_expr(&exprs[0])?);
                    } else {
                        return Err(RuntimeError::Custom("Multiple select fields not yet supported in runtime".to_string()));
                    }
                }
                SelectOrGroupClause::Group(_g) => {
                     return Err(RuntimeError::Custom("Group By not yet supported".to_string()));
                }
            }
        }
        
        self.env.pop_scope();
        Ok(Value::Array(results))
    }

    fn construct_xml(&mut self, node: &vybe_parser::ast::xml::XmlNode) -> Result<Value, RuntimeError> {
        use vybe_parser::ast::xml::*;
        match node {
            XmlNode::Element(el) => {
                let name = if let Some(prefix) = &el.name.prefix {
                    format!("{}:{}", prefix, el.name.local)
                } else {
                    el.name.local.clone()
                };
                
                let mut content = Vec::new();
                
                // Attributes creation
                // Note: builtins::xml::create_xattribute takes &[Value] usually [name, value]
                for attr in &el.attributes {
                    let attr_name = if let Some(prefix) = &attr.name.prefix {
                        format!("{}:{}", prefix, attr.name.local)
                    } else {
                        attr.name.local.clone()
                    };
                    
                    let mut attr_val_str = String::new();
                    // Attribute value can be composite
                    for part in &attr.value {
                         match part {
                             XmlNode::Text(s) => attr_val_str.push_str(s),
                             XmlNode::EmbeddedExpression(expr) => {
                                 let v = self.evaluate_expr(expr)?;
                                 attr_val_str.push_str(&v.as_string());
                             }
                             _ => {}
                         }
                    }
                    
                    let attr_obj = crate::builtins::xml::create_xattribute(&[
                        Value::String(attr_name),
                        Value::String(attr_val_str)
                    ]);
                    // Attributes are part of content/children in create_xelement for now or need separate handling?
                    // create_xelement description says "content can be: string value, XAttribute, XElement..."
                    content.push(attr_obj);
                }
                
                // Children
                for child in &el.children {
                    match child {
                        XmlNode::Text(s) => content.push(Value::String(s.clone())),
                        XmlNode::EmbeddedExpression(expr) => {
                             let v = self.evaluate_expr(expr)?;
                             content.push(v);
                        }
                        XmlNode::Element(_) => {
                            content.push(self.construct_xml(child)?);
                        }
                        XmlNode::Comment(s) => {
                            content.push(crate::builtins::xml::create_xcomment(&[Value::String(s.clone())]));
                        }
                        _ => {}
                    }
                }
                
                let mut args = vec![Value::String(name)];
                args.extend(content);
                
                let result = crate::builtins::xml::create_xelement(&args);
                Ok(result)
            }
            XmlNode::Text(s) => Ok(Value::String(s.clone())),
            XmlNode::EmbeddedExpression(expr) => self.evaluate_expr(expr),
            XmlNode::Comment(s) => Ok(crate::builtins::xml::create_xcomment(&[Value::String(s.clone())])),
            XmlNode::CData(s) => Ok(Value::String(s.clone())), // Treat CDATA as text for now
        }
    }
}
