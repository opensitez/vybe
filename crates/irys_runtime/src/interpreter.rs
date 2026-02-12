use crate::builtins::*;
use crate::environment::Environment;
use crate::evaluator::{evaluate, values_equal, value_in_range, compare_values};
use crate::data_access::DataAccessManager;
use crate::event_system::EventSystem;
use crate::value::{ExitType, RuntimeError, Value, ObjectData};
use crate::EventData;
use std::collections::{HashMap, VecDeque};
use std::rc::Rc;
use std::cell::RefCell;
use std::sync::mpsc;
use irys_parser::{CaseCondition, Declaration, Expression, FunctionDecl, Identifier, Program, Statement, SubDecl};

pub struct Interpreter {
    pub env: Environment,
    pub functions: HashMap<String, FunctionDecl>,
    pub subs: HashMap<String, SubDecl>,
    pub classes: HashMap<String, irys_parser::ClassDecl>,
    pub events: EventSystem,
    pub side_effects: VecDeque<crate::RuntimeSideEffect>,
    current_module: Option<String>, // Track which form/module is currently executing
    current_object: Option<Rc<RefCell<crate::value::ObjectData>>>,
    with_object: Option<Value>,
    pub file_handles: HashMap<i32, crate::file_io::FileHandle>,
    pub resources: HashMap<String, String>,
    pub resource_entries: Vec<crate::ResourceEntry>,
    pub command_line_args: Vec<String>,
    /// Optional channel for sending console output to the UI (interactive console mode).
    pub console_tx: Option<mpsc::Sender<crate::ConsoleMessage>>,
    /// Optional channel for receiving console input from the UI (Console.ReadLine).
    pub console_input_rx: Option<mpsc::Receiver<String>>,
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
            resources: HashMap::new(),
            resource_entries: Vec::new(),
            command_line_args: Vec::new(),
            console_tx: None,
            console_input_rx: None,
        };
        interp.register_builtin_constants();
        interp.init_namespaces();
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
        system_fields.insert("io".to_string(), io_obj);
        system_fields.insert("console".to_string(), console_obj.clone());
        system_fields.insert("math".to_string(), math_obj.clone());
        // System.DBNull
        let mut dbnull_fields = HashMap::new();
        dbnull_fields.insert("value".to_string(), Value::Nothing);
        let dbnull_obj = Value::Object(Rc::new(RefCell::new(ObjectData {
            class_name: "DBNull".to_string(), fields: dbnull_fields,
        })));
        system_fields.insert("dbnull".to_string(), dbnull_obj.clone());
        
        let system_obj_data = ObjectData {
            class_name: "Namespace".to_string(),
            fields: system_fields,
        };
        let system_obj = Value::Object(Rc::new(RefCell::new(system_obj_data)));

        // Register "System" in env
        self.env.define("system", system_obj);
        
        // Also register Console and Math globally for convenience (like implicit Imports System)
        self.env.define("console", console_obj);
        self.env.define("math", math_obj);
        // Also register DBNull globally
        self.env.define("dbnull", dbnull_obj);
        // Register ConsoleColor enum globally
        self.env.define("consolecolor", cc_obj);
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

    /// Send console output through the channel (with colors) or the side-effects queue.
    fn send_console_output(&mut self, text: String) {
        let (fg, bg) = self.get_console_colors();
        if let Some(tx) = &self.console_tx {
            let _ = tx.send(crate::ConsoleMessage::Output { text, fg, bg });
        } else {
            // CLI mode: only wrap in ANSI escape codes if colors differ from defaults
            if fg != 7 || bg != 0 {
                let ansi_fg = Self::console_color_to_ansi_fg(fg);
                let ansi_bg = Self::console_color_to_ansi_bg(bg);
                let colored = format!("{}{}{}\x1b[0m", ansi_fg, ansi_bg, text);
                self.side_effects.push_back(crate::RuntimeSideEffect::ConsoleOutput(colored));
            } else {
                self.side_effects.push_back(crate::RuntimeSideEffect::ConsoleOutput(text));
            }
        }
    }

    /// Send debug output (always uses default colors).
    fn send_debug_output(&mut self, text: String) {
        if let Some(tx) = &self.console_tx {
            let _ = tx.send(crate::ConsoleMessage::Output { text, fg: 7, bg: 0 });
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

    /// Create a System.Windows.Forms.MouseEventArgs object.
    pub fn make_mouse_event_args(button: i32, clicks: i32, x: i32, y: i32, delta: i32) -> Value {
        let mut fields = std::collections::HashMap::new();
        fields.insert("__type".to_string(), Value::String("MouseEventArgs".to_string()));
        fields.insert("button".to_string(), Value::Integer(button));
        fields.insert("clicks".to_string(), Value::Integer(clicks));
        fields.insert("x".to_string(), Value::Integer(x));
        fields.insert("y".to_string(), Value::Integer(y));
        fields.insert("delta".to_string(), Value::Integer(delta));
        fields.insert("location".to_string(), Value::String(format!("{{X={},Y={}}}", x, y)));
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

    fn register_builtin_constants(&mut self) {
        self.env.define_const("vbcrlf", Value::String("\r\n".to_string()));
        self.env.define_const("vbnewline", Value::String("\r\n".to_string()));
        self.env.define_const("vbtab", Value::String("\t".to_string()));
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
    /// These are the arguments *after* the project file on the irys command line.
    /// We prepend the program name (like .NET's Environment.GetCommandLineArgs).
    pub fn set_command_line_args(&mut self, args: Vec<String>) {
        let mut full = vec!["irys".to_string()];
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
                        existing.is_partial = existing.is_partial || class_decl.is_partial;
                    } else {
                        // Replace only when both are non-partial
                        *existing = class_decl.clone();
                    }
                } else {
                    self.classes.insert(key.clone(), class_decl.clone());
                }

                // Register class methods as subs so they can be called by event system
                for method in &class_decl.methods {
                    match method {
                        irys_parser::ast::MethodDecl::Sub(sub_decl) => {
                            let sub_key = format!("{}.{}", key, sub_decl.name.as_str().to_lowercase());
                            self.subs.insert(sub_key, sub_decl.clone());
                        }
                        irys_parser::ast::MethodDecl::Function(func_decl) => {
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
        }
    }

    // Helper to collect all fields including inherited ones
    fn collect_fields(&mut self, class_name: &str) -> HashMap<String, Value> {
        let mut fields = HashMap::new();
        
        // 1. Get base class fields first (if any)
        if let Some(cls) = self.classes.get(&class_name.to_lowercase()).cloned() {
             if let Some(parent_type) = &cls.inherits {
                 // Resolve parent type to string
                 // VBType::Custom(name) usually
                 let parent_name = match parent_type {
                     irys_parser::VBType::Custom(n) => Some(n.clone()),
                     irys_parser::VBType::Object => None, // Object has no fields
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
                             irys_parser::VBType::Integer => Value::Integer(0),
                             irys_parser::VBType::Long => Value::Long(0),
                             irys_parser::VBType::Single => Value::Single(0.0),
                             irys_parser::VBType::Double => Value::Double(0.0),
                             irys_parser::VBType::String => Value::String("".to_string()),
                             irys_parser::VBType::Boolean => Value::Boolean(false),
                             irys_parser::VBType::Custom(s) => {
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
                                    s_lower == "backgroundworker" ||
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
                                    s_lower == "printdocument" ||
                                    s_lower == "printpreviewcontrol" ||
                                    s_lower == "printpreviewdialog" ||
                                    s_lower == "pagesetupdialog" ||
                                    s_lower == "colordialog" ||
                                    s_lower == "fontdialog" ||
                                    s_lower == "folderbrowserdialog" ||
                                    s_lower == "opendialog" ||
                                    s_lower == "savefiledialog"
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
    fn find_method(&self, class_name: &str, method_name: &str) -> Option<irys_parser::ast::decl::MethodDecl> {
        let key = class_name.to_lowercase();
        if let Some(cls) = self.classes.get(&key) {
            // Check current class
            for method in &cls.methods {
                let m_name = match method {
                    irys_parser::ast::decl::MethodDecl::Sub(s) => &s.name,
                    irys_parser::ast::decl::MethodDecl::Function(f) => &f.name,
                };
                if m_name.as_str().eq_ignore_ascii_case(method_name) {
                    return Some(method.clone());
                }
            }
            
            // Check base class
             if let Some(parent_type) = &cls.inherits {
                 let parent_name = match parent_type {
                     irys_parser::VBType::Custom(n) => Some(n.clone()),
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
    fn find_property(&self, class_name: &str, prop_name: &str) -> Option<irys_parser::ast::decl::PropertyDecl> {
        let key = class_name.to_lowercase();
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
                     irys_parser::VBType::Custom(n) => Some(n.clone()),
                     _ => None,
                 };
                 if let Some(p_name) = parent_name {
                     return self.find_property(&p_name, prop_name);
                 }
             }
        }
        None
    }


    pub fn execute(&mut self, stmt: &Statement) -> Result<(), RuntimeError> {
        match stmt {
            Statement::Dim(decl) => {
                if let Some(bounds) = &decl.array_bounds {
                    // Array declaration with bounds: Dim arr(10) As Integer
                    let size = (self.evaluate_expr(&bounds[0])?.as_integer()? + 1) as usize; // VB arrays are 0-based but size is bound+1
                    let default_val = default_value_for_type("", &decl.var_type);
                    let arr = Value::Array(vec![default_val; size]);
                    self.env.define(decl.name.as_str(), arr);
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
                            // Store other BindingSource properties directly
                            if member_lower == "datamember" || member_lower == "filter" || member_lower == "sort" || member_lower == "position" {
                                obj_ref.borrow_mut().fields.insert(member_lower.clone(), val.clone());
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
                                "webbrowser" | "errorprovider" | "tooltip" | "backgroundworker" |
                                "datagridview" | "bindingnavigator" | "bindingsource" |
                                "flowlayoutpanel" | "tablelayoutpanel" | "splitcontainer" |
                                "maskedtextbox" | "domainupdown" | "contextmenustrip"
                            );
                            if is_winforms {
                                Value::String(prop_name.clone())
                            } else {
                                val.clone()
                            }
                        } else {
                            val.clone()
                        };
                        obj_ref.borrow_mut().fields.insert(member_lower.clone(), store_val);

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
                let index = self.evaluate_expr(&indices[0])?.as_integer()? as usize;
                let array_lower = array.as_str().to_lowercase();

                // Get the array, modify it, and store it back
                // Try current object fields first
                if let Some(obj_rc) = &self.current_object {
                    let mut obj = obj_rc.borrow_mut();
                    if let Some(arr_val) = obj.fields.get_mut(&array_lower) {
                        arr_val.set_array_element(index, val.clone())?;
                        return Ok(());
                    }
                }

                // Fallback to environment
                let mut arr = self.env.get(array.as_str())?;
                arr.set_array_element(index, val)?;
                self.env.set(array.as_str(), arr)?;
                Ok(())
            }

            Statement::ReDim { preserve, array, bounds } => {
                let new_size = (self.evaluate_expr(&bounds[0])?.as_integer()? + 1) as usize;

                if *preserve {
                    // Get existing array and resize preserving data
                    let mut arr = self.env.get(array.as_str())?;
                    if let Value::Array(ref mut vec) = arr {
                        vec.resize(new_size, Value::Nothing);
                        self.env.set(array.as_str(), arr)?;
                    } else {
                        return Err(RuntimeError::Custom(format!("{} is not an array", array.as_str())));
                    }
                } else {
                    // Create new array
                    let new_arr = Value::Array(vec![Value::Nothing; new_size]);
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
                            Err(RuntimeError::Continue(irys_parser::ast::stmt::ContinueType::For)) => break,
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
                            
                            Err(RuntimeError::Continue(irys_parser::ast::stmt::ContinueType::While)) => break,
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
                use irys_parser::LoopConditionType;

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
                            Err(RuntimeError::Continue(irys_parser::ast::stmt::ContinueType::Do)) => break,
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
                                 let when_match = if let Some(expr) = &catch.when_clause {
                                      self.evaluate_expr(expr)?.is_truthy()
                                 } else {
                                      true
                                 };
                                 
                                 if when_match {
                                     if let Some((name, _)) = &catch.variable {
                                          let mut ex_fields = std::collections::HashMap::new();
                                          ex_fields.insert("message".to_string(), Value::String(ex_msg.clone()));
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
                                     
                                     flow_result = self.execute_block(&catch.body);
                                     handled = true;
                                     break;
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
                self.call_procedure(name, arguments)?;
                Ok(())
            }

            Statement::ExpressionStatement(expr) => {
                self.evaluate_expr(expr)?;
                Ok(())
            }

            Statement::ForEach { variable, collection, body } => {
                let coll_val = self.evaluate_expr(collection)?;
                match coll_val {
                    Value::Array(items) => {
                        self.env.define(variable.as_str(), Value::Nothing);
                        for item in &items {
                            self.env.set(variable.as_str(), item.clone())?;
                            let mut should_exit = false;
                            for s in body {
                                match self.execute(s) {
                                    Err(RuntimeError::Exit(ExitType::For)) => { should_exit = true; break; }
                                    Err(RuntimeError::Continue(irys_parser::ast::stmt::ContinueType::For)) => break,
                                    Err(e) => return Err(e),
                                    Ok(()) => {}
                                }
                            }
                            if should_exit { break; }
                        }
                    }
                    Value::Collection(coll_rc) => {
                        // Support Collection/ArrayList iteration
                        let coll = coll_rc.borrow();
                        self.env.define(variable.as_str(), Value::Nothing);
                        for item in &coll.items {
                            self.env.set(variable.as_str(), item.clone())?;
                            let mut should_exit = false;
                            for s in body {
                                match self.execute(s) {
                                    Err(RuntimeError::Exit(ExitType::For)) => { should_exit = true; break; }
                                    Err(RuntimeError::Continue(irys_parser::ast::stmt::ContinueType::For)) => break,
                                    Err(e) => return Err(e),
                                    Ok(()) => {}
                                }
                            }
                            if should_exit { break; }
                        }
                    }
                    Value::String(s) => {
                        self.env.define(variable.as_str(), Value::String(String::new()));
                        for ch in s.chars() {
                            self.env.set(variable.as_str(), Value::String(ch.to_string()))?;
                            let mut should_exit = false;
                            for stmt in body {
                                match self.execute(stmt) {
                                    Err(RuntimeError::Exit(ExitType::For)) => { should_exit = true; break; }
                                    Err(RuntimeError::Continue(irys_parser::ast::stmt::ContinueType::For)) => break,
                                    Err(e) => return Err(e),
                                    Ok(()) => {}
                                }
                            }
                            if should_exit { break; }
                        }
                    }
                    _ => return Err(RuntimeError::Custom("For Each requires an array, collection, or string".to_string())),
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
                        if let irys_parser::ast::decl::MethodDecl::Sub(s) = method {
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
                           // Array access via Call syntax
                           if args.len() != 1 {
                               return Err(RuntimeError::Custom("Array index must be 1 dimension".to_string()));
                           }
                           let index = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                           return arr.get(index).cloned().ok_or_else(|| RuntimeError::Custom("Array index out of bounds".to_string()));
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
                self.call_method(obj, method, args)
            }
            Expression::ArrayAccess(array, indices) => {
                let arr = self.env.get(array.as_str())?;
                let index = self.evaluate_expr(&indices[0])?.as_integer()? as usize;
                arr.get_array_element(index)
            }
            Expression::ArrayLiteral(elements) => {
                let vals: Result<Vec<Value>, RuntimeError> = elements
                    .iter()
                    .map(|e| self.evaluate_expr(e))
                    .collect();
                Ok(Value::Array(vals?))
            }
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
            Expression::WithTarget => {
                if let Some(val) = &self.with_object {
                    Ok(val.clone())
                } else {
                    Err(RuntimeError::Custom("'.' used outside of With block".to_string()))
                }
            }
            Expression::New(class_id, ctor_args) => {
                let class_name = class_id.as_str().to_lowercase();

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
                    fields.insert("connected".to_string(), Value::Boolean(false));
                    fields.insert("receivebuffersize".to_string(), Value::Integer(8192));
                    fields.insert("sendbuffersize".to_string(), Value::Integer(8192));
                    fields.insert("__socket_id".to_string(), Value::Long(0));
                    fields.insert("__recv_buffer".to_string(), Value::Array(Vec::new()));
                    // If host and port provided, auto-connect
                    if !host.is_empty() && port > 0 {
                        fields.insert("connected".to_string(), Value::Boolean(true));
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
                    fields.insert("__recv_buffer".to_string(), Value::Array(Vec::new()));
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

                if class_name.starts_with("system.windows.forms.")
                    && class_name != "system.windows.forms.bindingsource"
                    && class_name != "system.windows.forms.bindingnavigator"
                {
                    // Return the base name as a proxy for controls (e.g. "Button", "Label")
                    let base_name = class_id.as_str().split('.').last().unwrap_or(class_id.as_str()).to_string();
                    return Ok(Value::String(base_name));
                }

                if class_name.starts_with("system.drawing.") 
                    || class_name.starts_with("system.componentmodel.") {
                    return Ok(Value::Nothing);
                }

                if class_name == "system.collections.arraylist" || class_name == "arraylist" {
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

                // System.IO.StreamReader
                if class_name == "streamreader" || class_name == "system.io.streamreader" {
                    let path = if !ctor_args.is_empty() {
                        self.evaluate_expr(&ctor_args[0])?.as_string()
                    } else {
                        return Err(RuntimeError::Custom("StreamReader requires a file path".to_string()));
                    };
                    let content = std::fs::read_to_string(&path)
                        .map_err(|e| RuntimeError::Custom(format!("StreamReader: {}", e)))?;
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("StreamReader".to_string()));
                    fields.insert("__content".to_string(), Value::String(content));
                    fields.insert("__position".to_string(), Value::Integer(0));
                    fields.insert("__closed".to_string(), Value::Boolean(false));
                    let obj = crate::value::ObjectData { class_name: "StreamReader".to_string(), fields };
                    return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                }

                // System.IO.StreamWriter
                if class_name == "streamwriter" || class_name == "system.io.streamwriter" {
                    let path = if !ctor_args.is_empty() {
                        self.evaluate_expr(&ctor_args[0])?.as_string()
                    } else {
                        return Err(RuntimeError::Custom("StreamWriter requires a file path".to_string()));
                    };
                    let append = if ctor_args.len() >= 2 {
                        self.evaluate_expr(&ctor_args[1])?.as_bool()?
                    } else {
                        false
                    };
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("__type".to_string(), Value::String("StreamWriter".to_string()));
                    fields.insert("__path".to_string(), Value::String(path));
                    fields.insert("__buffer".to_string(), Value::String(String::new()));
                    fields.insert("__append".to_string(), Value::Boolean(append));
                    fields.insert("__closed".to_string(), Value::Boolean(false));
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

                if let Some(class_decl) = self.classes.get(&class_name).cloned() {
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
                        match method {
                            irys_parser::ast::decl::MethodDecl::Sub(s) => {
                                self.call_user_sub(&s, &arg_values, Some(obj_ref.clone()))?;
                            }
                            irys_parser::ast::decl::MethodDecl::Function(f) => {
                                self.call_user_function(&f, &arg_values, Some(obj_ref.clone()))?;
                            }
                        }
                    } else {
                        // No Sub New found: for WinForms classes (inheriting System.Windows.Forms.Form),
                        // auto-call InitializeComponent if it exists
                        let inherits_form = class_decl.inherits.as_ref().map_or(false, |t| {
                            match t {
                                irys_parser::VBType::Custom(n) => n.to_lowercase().contains("form"),
                                _ => false,
                            }
                        });
                        if inherits_form {
                            if let Some(init_method) = self.find_method(&class_name, "InitializeComponent") {
                                match init_method {
                                    irys_parser::ast::decl::MethodDecl::Sub(s) => {
                                        let _ = self.call_user_sub(&s, &[], Some(obj_ref.clone()));
                                    }
                                    irys_parser::ast::decl::MethodDecl::Function(f) => {
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
                    if member.as_str().eq_ignore_ascii_case("Count") {
                        return Ok(Value::Integer(col_rc.borrow().count()));
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
                        if let Some(val) = obj_data.fields.get(&member.as_str().to_lowercase()) {
                            return Ok(val.clone());
                        }
                    } // Drop borrow

                    // Handle WinForms infrastructure properties that don't exist as real fields
                    let member_lower = member.as_str().to_lowercase();
                    match member_lower.as_str() {
                        "controls" | "components" => return Ok(Value::Nothing),
                        "databindings" => {
                            // Return a DataBindings proxy for this control object
                            let mut flds = std::collections::HashMap::new();
                            flds.insert("__type".to_string(), Value::String("DataBindings".to_string()));
                            flds.insert("__parent".to_string(), Value::Object(obj_ref.clone()));
                            flds.insert("__parent_name".to_string(), {
                                let borrow = obj_ref.borrow();
                                borrow.fields.get("name")
                                    .cloned()
                                    .unwrap_or(Value::String(class_name_str.clone()))
                            });
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
                        return Ok(Value::Array(Vec::new()));
                    }
                    let key = format!("{}.{}", obj_name, member.as_str());
                    if let Ok(val) = self.env.get(&key) {
                        return Ok(val);
                    }
                    // For string proxy control properties not yet synced, return empty string
                    return Ok(Value::String(String::new()));
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
                    // If field not present, create it as Nothing to mimic VB's instance fields being default-initialized
                    drop(obj);
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

                // 5. Fallback: implicit function call without parentheses (e.g. "Now", "Date")
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
                    irys_parser::ast::decl::MethodDecl::Sub(s) => {
                        self.call_user_sub_exprs(&s, args, Some(obj_rc.clone()))?;
                        return Ok(Value::Nothing);
                    }
                    irys_parser::ast::decl::MethodDecl::Function(f) => {
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
        
        // Try 3: Global search - look for any function with matching unqualified name
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
        
        // Try 3: Global search for subs
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
                self.side_effects.push_back(crate::RuntimeSideEffect::MsgBox(msg));
                return Ok(Value::Integer(1)); // vbOK
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
            "xdocument.parse" | "xml.parse" => return xml_parse_fn(&arg_values),
            "xdocument.save" | "xml.save" => return xml_save_fn(&arg_values),

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

        // Handle WinForms designer no-op methods
        match method_name.as_str() {
            "suspendlayout" | "resumelayout" | "performlayout" => return Ok(Value::Nothing),
            _ => {}
        }

        // Handle Controls.Add - a common designer pattern
        if method_name == "add" {
            if let Expression::MemberAccess(_, member) = obj {
                if member.as_str().eq_ignore_ascii_case("Controls") {
                    return Ok(Value::Nothing);
                }
            }
        }

        // Evaluate object to check if it's a Collection or Dialog
        if let Ok(obj_val) = self.evaluate_expr(obj) {
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
                    // ToString with format
                    "tostring" => {
                        let fmt = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                        return Ok(Value::String(format_ole_date(*ole_val, &fmt)));
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

            // Handle StringBuilder methods
            if let Value::Object(obj_ref) = &obj_val {
                let type_name = obj_ref.borrow().fields.get("__type").and_then(|v| {
                    if let Value::String(s) = v { Some(s.clone()) } else { None }
                }).unwrap_or_default();
                if !type_name.is_empty() {
                    if type_name == "StringBuilder" {
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

                    // StreamReader instance methods
                    if type_name == "StreamReader" {
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

                    // StreamWriter instance methods
                    if type_name == "StreamWriter" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter()
                            .map(|arg| self.evaluate_expr(arg))
                            .collect();
                        let arg_values = arg_values?;
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
                                let replacement = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                                let args_for_fn = vec![Value::String(input), Value::String(full_pattern), Value::String(replacement)];
                                return crate::builtins::text_fns::regex_replace_fn(&args_for_fn);
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
                                obj_ref.borrow_mut().fields.insert("__host".to_string(), Value::String(host));
                                obj_ref.borrow_mut().fields.insert("__port".to_string(), Value::Integer(port));
                                obj_ref.borrow_mut().fields.insert("connected".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            "getstream" => {
                                // Return a NetworkStream-like object (MemoryStream proxy)
                                let mut fields = std::collections::HashMap::new();
                                fields.insert("__type".to_string(), Value::String("NetworkStream".to_string()));
                                fields.insert("__data".to_string(), Value::Array(Vec::new()));
                                fields.insert("canread".to_string(), Value::Boolean(true));
                                fields.insert("canwrite".to_string(), Value::Boolean(true));
                                fields.insert("position".to_string(), Value::Long(0));
                                fields.insert("length".to_string(), Value::Long(0));
                                fields.insert("dataavailable".to_string(), Value::Boolean(false));
                                let obj = crate::value::ObjectData { class_name: "NetworkStream".to_string(), fields };
                                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                            }
                            "close" | "dispose" => {
                                obj_ref.borrow_mut().fields.insert("connected".to_string(), Value::Boolean(false));
                                return Ok(Value::Nothing);
                            }
                            _ => {}
                        }
                    }

                    // ===== TcpListener instance methods =====
                    if type_name == "TcpListener" {
                        match method_name.as_str() {
                            "start" => {
                                obj_ref.borrow_mut().fields.insert("__active".to_string(), Value::Boolean(true));
                                return Ok(Value::Nothing);
                            }
                            "stop" => {
                                obj_ref.borrow_mut().fields.insert("__active".to_string(), Value::Boolean(false));
                                return Ok(Value::Nothing);
                            }
                            "accepttcpclient" => {
                                // Return a new TcpClient stub
                                let mut fields = std::collections::HashMap::new();
                                fields.insert("__type".to_string(), Value::String("TcpClient".to_string()));
                                fields.insert("connected".to_string(), Value::Boolean(true));
                                fields.insert("__host".to_string(), Value::String("127.0.0.1".to_string()));
                                fields.insert("__port".to_string(), Value::Integer(0));
                                fields.insert("receivebuffersize".to_string(), Value::Integer(8192));
                                fields.insert("sendbuffersize".to_string(), Value::Integer(8192));
                                fields.insert("__recv_buffer".to_string(), Value::Array(Vec::new()));
                                let obj = crate::value::ObjectData { class_name: "TcpClient".to_string(), fields };
                                return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                            }
                            "pending" => {
                                return Ok(Value::Boolean(false));
                            }
                            _ => {}
                        }
                    }

                    // ===== UdpClient instance methods =====
                    if type_name == "UdpClient" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                        let arg_values = arg_values?;
                        match method_name.as_str() {
                            "send" => {
                                // send(bytes, length, hostname, port)
                                let _length = arg_values.get(1).and_then(|v| v.as_integer().ok()).unwrap_or(0);
                                return Ok(Value::Integer(_length));
                            }
                            "receive" => {
                                // Return empty byte array (stub)
                                return Ok(Value::Array(Vec::new()));
                            }
                            "close" | "dispose" => { return Ok(Value::Nothing); }
                            _ => {}
                        }
                    }

                    // ===== SmtpClient instance methods =====
                    if type_name == "SmtpClient" {
                        let arg_values: Result<Vec<Value>, RuntimeError> = args.iter().map(|a| self.evaluate_expr(a)).collect();
                        let arg_values = arg_values?;
                        match method_name.as_str() {
                            "send" => {
                                // SmtpClient.Send(MailMessage) — use system mail command if available
                                if let Some(Value::Object(msg_ref)) = arg_values.get(0) {
                                    let msg = msg_ref.borrow();
                                    let from = msg.fields.get("from").map(|v| v.as_string()).unwrap_or_default();
                                    let to = msg.fields.get("to").map(|v| v.as_string()).unwrap_or_default();
                                    let subject = msg.fields.get("subject").map(|v| v.as_string()).unwrap_or_default();
                                    let body = msg.fields.get("body").map(|v| v.as_string()).unwrap_or_default();
                                    let host = obj_ref.borrow().fields.get("host").map(|v| v.as_string()).unwrap_or_default();
                                    let port = obj_ref.borrow().fields.get("port").and_then(|v| v.as_integer().ok()).unwrap_or(25);
                                    // Use curl to send email via SMTP
                                    let smtp_url = format!("smtp://{}:{}", host, port);
                                    let status = std::process::Command::new("curl")
                                        .args(&[
                                            "--mail-from", &from,
                                            "--mail-rcpt", &to,
                                            "--url", &smtp_url,
                                            "-T", "-",
                                        ])
                                        .stdin(std::process::Stdio::piped())
                                        .spawn()
                                        .and_then(|mut child| {
                                            if let Some(ref mut stdin) = child.stdin {
                                                use std::io::Write;
                                                let _ = write!(stdin, "From: {}\r\nTo: {}\r\nSubject: {}\r\n\r\n{}", from, to, subject, body);
                                            }
                                            child.wait()
                                        });
                                    match status {
                                        Ok(s) if s.success() => return Ok(Value::Nothing),
                                        _ => return Err(RuntimeError::Custom("SmtpClient.Send failed".to_string())),
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
                            let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                            let count = self.binding_source_row_count(&ds);
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
                            self.refresh_bindings(obj_ref, &ds, new_pos);
                            return Ok(Value::Nothing);
                        }
                        "moveprevious" => {
                            let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                            let count = self.binding_source_row_count(&ds);
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
                            self.refresh_bindings(obj_ref, &ds, new_pos);
                            return Ok(Value::Nothing);
                        }
                        "movefirst" => {
                            let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                            let count = self.binding_source_row_count(&ds);
                            obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Integer(0));
                            self.side_effects.push_back(crate::RuntimeSideEffect::BindingPositionChanged {
                                binding_source_name: bs_name.clone(),
                                position: 0,
                                count,
                            });
                            self.refresh_bindings(obj_ref, &ds, 0);
                            return Ok(Value::Nothing);
                        }
                        "movelast" => {
                            let ds = obj_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                            let count = self.binding_source_row_count(&ds);
                            let last = if count > 0 { count - 1 } else { 0 };
                            obj_ref.borrow_mut().fields.insert("position".to_string(), Value::Integer(last));
                            self.side_effects.push_back(crate::RuntimeSideEffect::BindingPositionChanged {
                                binding_source_name: bs_name.clone(),
                                position: last,
                                count,
                            });
                            self.refresh_bindings(obj_ref, &ds, last);
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
                                .map(|v| v.as_string()).unwrap_or_default();

                            // Store the binding info on the BindingSource if data_source is a BindingSource
                            if let Value::Object(ds_ref) = &data_source {
                                let ds_type = ds_ref.borrow().fields.get("__type")
                                    .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                                    .unwrap_or_default();
                                if ds_type == "BindingSource" {
                                    let mut ds_borrow = ds_ref.borrow_mut();
                                    if let Some(Value::Array(bindings)) = ds_borrow.fields.get_mut("__bindings") {
                                        // Store binding as "controlName|propertyName|dataMember"
                                        let binding_entry = format!("{}|{}|{}", parent_name, prop_name, data_member);
                                        bindings.push(Value::String(binding_entry));
                                    }
                                }
                            }

                            // Store binding info in environment for the control
                            let binding_key = format!("__binding_{}_{}", parent_name, prop_name);
                            let mut binding_fields = std::collections::HashMap::new();
                            binding_fields.insert("property".to_string(), Value::String(prop_name.clone()));
                            binding_fields.insert("datasource".to_string(), data_source.clone());
                            binding_fields.insert("datamember".to_string(), Value::String(data_member.clone()));
                            binding_fields.insert("controlname".to_string(), Value::String(parent_name.clone()));
                            let binding_obj = crate::value::ObjectData { class_name: "Binding".to_string(), fields: binding_fields };
                            self.env.define_global(&binding_key, Value::Object(std::rc::Rc::new(std::cell::RefCell::new(binding_obj))));

                            // Try to immediately sync — get current value from BindingSource
                            if let Value::Object(ds_ref) = &data_source {
                                let ds_type = ds_ref.borrow().fields.get("__type")
                                    .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                                    .unwrap_or_default();
                                if ds_type == "BindingSource" {
                                    let ds_val = ds_ref.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                                    let dm = ds_ref.borrow().fields.get("datamember")
                                        .map(|v| v.as_string()).unwrap_or_default();
                                    eprintln!("[DataBindings.Add] ctrl={} prop={} dm={} ds_val={}", parent_name, prop_name, data_member, match &ds_val { Value::Nothing => "Nothing", Value::Object(_) => "Object", Value::String(_) => "String", _ => "other" });
                                    Self::inject_select_from_data_member(&ds_val, &dm);
                                    let pos = ds_ref.borrow().fields.get("position")
                                        .and_then(|v| if let Value::Integer(i) = v { Some(*i) } else { None })
                                        .unwrap_or(0);
                                    let row = self.binding_source_get_row(&ds_val, pos);
                                    eprintln!("[DataBindings.Add] row={}", match &row { Value::Nothing => "Nothing", Value::Object(_) => "Object", _ => "other" });
                                    if let Value::Object(row_ref) = &row {
                                        let member_lower = data_member.to_lowercase();
                                        let cell_val = row_ref.borrow().fields.get(&member_lower)
                                            .cloned().unwrap_or(Value::String(String::new()));
                                        eprintln!("[DataBindings.Add] cell_val for '{}' = {:?}", member_lower, cell_val.as_string());
                                        self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                                            object: parent_name.clone(),
                                            property: prop_name.clone(),
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
                    "length" | "count" => {
                        return Ok(Value::Integer(arr.len() as i32));
                    }
                    _ => {}
                }
            }

            if let Value::Collection(col_rc) = &obj_val {
                 match method_name.as_str() {
                    "add" => {
                        let val = self.evaluate_expr(&args[0])?;
                        let idx = col_rc.borrow_mut().add(val);
                        return Ok(Value::Integer(idx));
                    }
                    "remove" => {
                        let val = self.evaluate_expr(&args[0])?;
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
                         let idx = self.evaluate_expr(&args[0])?.as_integer()? as usize;
                         return col_rc.borrow().item(idx);
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
                    "reverse" => {
                        col_rc.borrow_mut().items.reverse();
                        return Ok(Value::Nothing);
                    }
                    "indexof" => {
                        let val = self.evaluate_expr(&args[0])?;
                        let idx = col_rc.borrow().items.iter().position(|v| {
                            crate::evaluator::values_equal(v, &val)
                        }).map(|i| i as i32).unwrap_or(-1);
                        return Ok(Value::Integer(idx));
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
                        let idx = col_rc.borrow().items.iter().rposition(|v| {
                            crate::evaluator::values_equal(v, &val)
                        }).map(|i| i as i32).unwrap_or(-1);
                        return Ok(Value::Integer(idx));
                    }
                    _ => {} // Fall through to other handlers
                 }
            }

            // Queue methods
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
            _ => return Err(RuntimeError::Custom("Invalid object reference".to_string())),
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
            // Method call on Nothing - silently ignore (common in WinForms designer for Controls.Add etc.)
            if obj_val == Value::Nothing {
                return Ok(Value::Nothing);
            }
            // String proxy: the object is a control name string (WinForms pattern)
            // Property access like btn.Caption resolves to env key "btn0.Caption"
            if let Value::String(obj_name) = &obj_val {
                let key = format!("{}.{}", obj_name, method.as_str());
                if let Ok(val) = self.env.get(&key) {
                    return Ok(val);
                }
                // Not found in env - return empty string as default for control properties
                return Ok(Value::String(String::new()));
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
                } else if class_name_lower == "system.math" {
                    return self.dispatch_math_method(&method_name, &arg_values);
                }
                
                // Use helper to find method in hierarchy
                if let Some(method) = self.find_method(&class_name_str, &method_name) {
                     match method {
                         irys_parser::ast::decl::MethodDecl::Sub(s) => {
                             self.call_user_sub(&s, &arg_values, Some(obj_ref.clone()))?;
                             return Ok(Value::Nothing);
                         }
                         irys_parser::ast::decl::MethodDecl::Function(f) => {
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
                use std::io::Write;
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
                // Task.Run(action) — runs lambda synchronously (single-threaded interpreter)
                // and wraps result in a Task object
                let result = if let Some(lambda) = arg_values.get(0) {
                    self.call_lambda(lambda.clone(), &[])?
                } else {
                    Value::Nothing
                };
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
            "task.delay" | "system.threading.tasks.task.delay" => {
                let ms = arg_values.get(0).map(|v| v.as_integer().unwrap_or(0)).unwrap_or(0);
                std::thread::sleep(std::time::Duration::from_millis(ms.max(0) as u64));
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
                // Execute callback synchronously
                if let Some(lambda) = arg_values.get(0) {
                    let state = arg_values.get(1).cloned().unwrap_or(Value::Nothing);
                    self.call_lambda(lambda.clone(), &[state])?;
                }
                return Ok(Value::Boolean(true));
            }
            "threadpool.setminthreads" | "system.threading.threadpool.setminthreads" => {
                return Ok(Value::Boolean(true)); // No-op, always succeed
            }
            "threadpool.setmaxthreads" | "system.threading.threadpool.setmaxthreads" => {
                return Ok(Value::Boolean(true));
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
                let file = arg_values.get(0).map(|v| v.as_string()).unwrap_or_default();
                let process_args = arg_values.get(1).map(|v| v.as_string()).unwrap_or_default();
                let mut cmd = std::process::Command::new(&file);
                if !process_args.is_empty() {
                    // Split args by space (simple split)
                    for a in process_args.split_whitespace() {
                        cmd.arg(a);
                    }
                }
                match cmd.spawn() {
                    Ok(child) => {
                        // Return a Process object with Id
                        let mut fields = std::collections::HashMap::new();
                        fields.insert("__type".to_string(), Value::String("Process".to_string()));
                        fields.insert("id".to_string(), Value::Integer(child.id() as i32));
                        fields.insert("hasexited".to_string(), Value::Boolean(false));
                        let obj = crate::value::ObjectData { class_name: "Process".to_string(), fields };
                        return Ok(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
                    }
                    Err(e) => return Err(RuntimeError::Custom(format!("Process.Start failed: {}", e))),
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
                fields.insert("useragent".to_string(), Value::String("irys/1.0".to_string()));
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
                }
            }
        }

        // Handle form methods
        match method_name.as_str() {
            "show" => {
                // Show the form
                self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                    object: object_name.clone(),
                    property: "Visible".to_string(),
                    value: Value::Boolean(true),
                });
                Ok(Value::Nothing)
            }
            "hide" => {
                // Hide the form
                self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                    object: object_name.clone(),
                    property: "Visible".to_string(),
                    value: Value::Boolean(false),
                });
                Ok(Value::Nothing)
            }
            "move" => {
                // Move form: .Move(left, top, width, height)
                if arg_values.len() >= 2 {
                    self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                        object: object_name.clone(),
                        property: "Left".to_string(),
                        value: arg_values[0].clone(),
                    });
                    self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                        object: object_name.clone(),
                        property: "Top".to_string(),
                        value: arg_values[1].clone(),
                    });
                    if arg_values.len() >= 3 {
                        self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                            object: object_name.clone(),
                            property: "Width".to_string(),
                            value: arg_values[2].clone(),
                        });
                    }
                    if arg_values.len() >= 4 {
                        self.side_effects.push_back(crate::RuntimeSideEffect::PropertyChange {
                            object: object_name.clone(),
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
                        object: object_name.clone(),
                        property: "URL".to_string(),
                        value: arg_values[0].clone(),
                    });
                }
                Ok(Value::Nothing)
            }
            _ => Err(RuntimeError::Custom(format!("Unknown method: {}", method_name))),
        }
    }

    // Generic helper for subs
    fn call_user_sub_impl(&mut self, sub: &SubDecl, args: Option<&[Value]>, arg_exprs: Option<&[Expression]>, context: Option<Rc<RefCell<ObjectData>>>) -> Result<Value, RuntimeError> {
        // Push new scope
        self.env.push_scope();

        // Save previous object context and set new one
        let prev_object = self.current_object.take();
        self.current_object = context;

        let mut byref_writebacks = Vec::new();

        // Bind parameters
        for (i, param) in sub.parameters.iter().enumerate() {
            let mut val = Value::Nothing;
            
            if let Some(values) = args {
                if i < values.len() {
                    val = values[i].clone();
                }
            } else if let Some(exprs) = arg_exprs {
                if i < exprs.len() {
                    val = self.evaluate_expr(&exprs[i])?;
                    let is_byref = param.pass_type == irys_parser::ast::decl::ParameterPassType::ByRef;
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

        // Execute body
        for stmt in &sub.body {
            match self.execute(stmt) {
                Err(RuntimeError::Exit(ExitType::Sub)) => break,
                Err(e) => {
                    self.env.pop_scope();
                    self.current_object = prev_object; // Restore context
                    return Err(e);
                }
                Ok(_) => {}
            }
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

        // Pop scope
        self.env.pop_scope();
        self.current_object = prev_object;
        
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

        let mut byref_writebacks = Vec::new();

        for (i, param) in func.parameters.iter().enumerate() {
            let mut val = Value::Nothing;
             if let Some(values) = args {
                if i < values.len() {
                    val = values[i].clone();
                }
            } else if let Some(exprs) = arg_exprs {
                if i < exprs.len() {
                    val = self.evaluate_expr(&exprs[i])?;
                    let is_byref = param.pass_type == irys_parser::ast::decl::ParameterPassType::ByRef;
                    if is_byref {
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

        self.env.define(func.name.as_str(), Value::Nothing);
        let mut explicit_return: Option<Value> = None;

        let mut result = Value::Nothing;
        for stmt in &func.body {
            match self.execute(stmt) {
                Err(RuntimeError::Exit(ExitType::Function)) => break,
                Err(RuntimeError::Return(val)) => {
                    if let Some(v) = val {
                        explicit_return = Some(v);
                    }
                    break;
                }
                Err(e) => {
                    self.env.pop_scope();
                    self.current_object = prev_object;
                    return Err(e);
                }
                Ok(_) => {}
            }
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

        self.env.pop_scope();
        self.current_object = prev_object;

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
            let class_name = obj_ref.borrow().class_name.clone();
            if let Some(method) = self.find_method(&class_name, method_name) {
                match method {
                    irys_parser::ast::decl::MethodDecl::Sub(s) => {
                        self.call_user_sub(&s, args, Some(obj_ref.clone()))?;
                        return Ok(());
                    }
                    irys_parser::ast::decl::MethodDecl::Function(f) => {
                        let _ = self.call_user_function(&f, args, Some(obj_ref.clone()))?;
                        return Ok(());
                    }
                }
            }
            Err(RuntimeError::UndefinedFunction(method_name.to_string()))
        } else {
            Err(RuntimeError::TypeError { expected: "Object".to_string(), got: format!("{:?}", instance_val) })
        }
    }

    /// Find a class method that has a Handles clause matching the given control.event pattern.
    /// For example, find_handles_method("form1", "btn0", "Click") finds a method with `Handles btn0.Click`.
    /// For Me.Load, use control_name="Me" and event_name="Load".
    pub fn find_handles_method(&self, class_name: &str, control_name: &str, event_name: &str) -> Option<String> {
        let key = class_name.to_lowercase();
        let target = format!("{}.{}", control_name, event_name).to_lowercase();
        if let Some(cls) = self.classes.get(&key) {
            for method in &cls.methods {
                if let irys_parser::ast::decl::MethodDecl::Sub(s) = method {
                    if let Some(handles) = &s.handles {
                        for h in handles {
                            if h.to_lowercase() == target {
                                return Some(s.name.as_str().to_string());
                            }
                        }
                    }
                }
            }
        }
        None
    }

    /// Get all Handles clause mappings for a class as (control, event) -> method_name.
    pub fn get_handles_map(&self, class_name: &str) -> HashMap<(String, String), String> {
        let mut map = HashMap::new();
        let key = class_name.to_lowercase();
        if let Some(cls) = self.classes.get(&key) {
            for method in &cls.methods {
                if let irys_parser::ast::decl::MethodDecl::Sub(s) = method {
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

    pub fn trigger_event(&mut self, control_name: &str, event_type: irys_forms::EventType, index: Option<i32>) -> Result<(), RuntimeError> {
        let handler_name = self.events.get_handler(control_name, &event_type).map(|s| s.to_string());
        if let Some(handler_name) = handler_name {
            // Check if it's a valid sub before calling
            if self.subs.contains_key(&handler_name.to_lowercase()) {
                let args: Vec<Value> = if let Some(idx) = index {
                    // VB6 control array: pass Index As Integer
                    vec![Value::Integer(idx)]
                } else {
                    // .NET style: pass (sender As Object, e As EventArgs)
                    self.make_event_handler_args(control_name, event_type.as_str())
                };
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

        eprintln!("[refresh_bindings] pos={} bindings={:?}", position, bindings);
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
                 let ndt = chrono::NaiveDateTime::from_timestamp_opt(secs, 0).unwrap_or_default();
                 Ok(Value::Date(date_to_ole(ndt)))
            }
            "getcreationtime" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 let meta = std::fs::metadata(&path).map_err(|e| RuntimeError::Custom(format!("Error accessing file: {}", e)))?;
                 let created = meta.created().unwrap_or(std::time::SystemTime::UNIX_EPOCH);
                 let secs = created.duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_secs() as i64;
                 let ndt = chrono::NaiveDateTime::from_timestamp_opt(secs, 0).unwrap_or_default();
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
                    // Fallback (CLI): read from real stdin
                    let mut line = String::new();
                    match std::io::stdin().read_line(&mut line) {
                        Ok(_) => {
                            if line.ends_with('\n') { line.pop(); }
                            if line.ends_with('\r') { line.pop(); }
                            Ok(Value::String(line))
                        }
                        Err(_) => Ok(Value::String(String::new())),
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
            let prev_env = self.env.clone();
            
            // Switch to captured environment (Snapshot)
            self.env = env.borrow().clone();
            self.env.push_scope();
            
            // Bind arguments
            for (param, arg) in params.iter().zip(args.iter()) {
                self.env.define(param.name.as_str(), arg.clone());
            }
            
            let result = match &*body {
                irys_parser::ast::expr::LambdaBody::Expression(expr) => {
                    self.evaluate_expr(expr)
                }
                irys_parser::ast::expr::LambdaBody::Statement(stmt) => {
                    self.execute(stmt)?;
                    Ok(Value::Nothing)
                }
            };
            
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

fn default_value_for_type(_name: &str, var_type: &Option<irys_parser::VBType>) -> Value {
    match var_type {
        Some(irys_parser::VBType::Integer) => Value::Integer(0),
        Some(irys_parser::VBType::Long) => Value::Long(0),
        Some(irys_parser::VBType::Single) => Value::Single(0.0),
        Some(irys_parser::VBType::Double) => Value::Double(0.0),
        Some(irys_parser::VBType::String) => Value::String(String::new()),
        Some(irys_parser::VBType::Boolean) => Value::Boolean(false),
        Some(irys_parser::VBType::Variant) => Value::Nothing,
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
