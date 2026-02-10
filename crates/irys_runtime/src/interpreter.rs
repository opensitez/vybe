use crate::builtins::*;
use crate::environment::Environment;
use crate::evaluator::{evaluate, values_equal, value_in_range, compare_values};
use crate::event_system::EventSystem;
use crate::value::{ExitType, RuntimeError, Value, ObjectData};
use std::collections::{HashMap, VecDeque};
use std::rc::Rc;
use std::cell::RefCell;
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
    pub command_line_args: Vec<String>,
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
            command_line_args: Vec::new(),
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

        // Create System.Console object
        let console_obj_data = ObjectData {
            class_name: "System.Console".to_string(),
            fields: HashMap::new(),
        };
        let console_obj = Value::Object(Rc::new(RefCell::new(console_obj_data)));

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
    }

    pub fn register_resources(&mut self, resources: HashMap<String, String>) {
        self.resources = resources.clone();

        // Create Resources object
        let mut res_fields = HashMap::new();
        for (key, val) in resources {
            res_fields.insert(key, Value::String(val));
        }

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
                        obj_ref.borrow_mut().fields.insert(member_lower.clone(), val.clone());

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
                        for catch in catches {
                             // Check variable type match if present
                             let type_match = if let Some((_, Some(_type_name))) = &catch.variable {
                                 // Simple type check? For now catch all or check error string?
                                 // We don't have typed exceptions yet, mostly checks.
                                 // For now, assume generic Catch catches everything.
                                 // If specialized, we might check error type.
                                 true 
                             } else {
                                 true // Catch All
                             };
                             
                             if type_match {
                                 // Check When clause
                                 let when_match = if let Some(expr) = &catch.when_clause {
                                      self.evaluate_expr(expr)?.is_truthy()
                                 } else {
                                      true
                                 };
                                 
                                 if when_match {
                                     // Execute Catch Body
                                     if let Some((name, _)) = &catch.variable {
                                          // Define exception variable
                                          // For now, simple string message
                                          let msg = format!("{}", e);
                                          // Define local variable?
                                          // We need internal scope or just use current scope?
                                          // VB.NET catch variable is local to catch block usually.
                                          // But our scope is function-level usually.
                                          // We can set it in env.
                                          self.env.set(name.as_str(), Value::String(msg))?;
                                     }
                                     
                                     flow_result = self.execute_block(&catch.body);
                                     handled = true;
                                     break;
                                 }
                             }
                        }
                        
                        if !handled {
                            // If not handled, keep original error
                        } else {
                             // If handled, flow_result is now result of catch block (Ok or new Err)
                        }
                    }
                }
                
                // Finally block
                if let Some(final_stmts) = finally {
                     let final_res = self.execute_block(final_stmts);
                     // If finally errors (or returns), it overrides previous result
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
                if let Ok(val) = self.evaluate_expr(&Expression::Variable(name.clone())) {
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
                // In Phase 1 "Simulation", Await simply evaluates the operand
                // If it were a real Task, we would wait for it.
                // Here we just return the value (identity behavior for now)
                self.evaluate_expr(operand)
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

                if class_name.starts_with("system.windows.forms.") {
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

                if let Value::Object(obj_ref) = &obj_val {
                    let class_name_str;
                    {
                        let obj_data = obj_ref.borrow();
                        class_name_str = obj_data.class_name.clone();
                        
                        // 1. Check if it's a field in the map (case-insensitive)
                        if let Some(val) = obj_data.fields.get(&member.as_str().to_lowercase()) {
                            return Ok(val.clone());
                        }
                    } // Drop borrow

                    // Handle WinForms infrastructure properties that don't exist as real fields
                    let member_lower = member.as_str().to_lowercase();
                    match member_lower.as_str() {
                        "controls" | "components" => return Ok(Value::Nothing),
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
                // Debug: trace variable resolution for controls
                // eprintln!("resolve var: {}", var_name);
                
                // 1. Check local scopes (everything except index 0 which is global)
                // We don't have a get_local yet, but we can search or just use env.get if we know it's not global
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
            "oct" | "oct$" => return oct_fn(&arg_values),
            "hex" | "hex$" => return hex_fn(&arg_values),

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
            "debug.print" | "console.writeline" | "console.write" => {
                let msg = arg_values.iter().map(|v| v.as_string()).collect::<Vec<_>>().join(" ");
                let with_newline = name_str.ends_with("writeline");
                let final_msg = if with_newline { format!("{}\n", msg) } else { msg };
                self.side_effects.push_back(crate::RuntimeSideEffect::ConsoleOutput(final_msg));
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
            // Handle StringBuilder methods
            if let Value::Object(obj_ref) = &obj_val {
                if let Some(Value::String(type_name)) = obj_ref.borrow().fields.get("__type") {
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
                self.side_effects.push_back(crate::RuntimeSideEffect::MsgBox(format!("[Debug] {}", msg)));
                return Ok(Value::Nothing);
            }

            // ---- Convert class ----
            "convert.todatetime" => {
                return crate::builtins::cdate_fn(&arg_values);
            }
            "convert.toint32" | "convert.toint16" | "convert.toint64" => {
                return crate::builtins::cint_fn(&arg_values);
            }
            "convert.todouble" | "convert.tosingle" | "convert.todecimal" => {
                return crate::builtins::cdbl_fn(&arg_values);
            }
            "convert.tostring" => {
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

            _ => {}
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
                    vec![Value::Integer(idx)]
                } else {
                    vec![]
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
            _ => Err(RuntimeError::UndefinedFunction(format!("System.IO.File.{}", method_name)))
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
                 Ok(Value::String(std::path::Path::new(&path).extension().unwrap_or_default().to_string_lossy().to_string()))
            }
            "getdirectoryname" => {
                 let path = args.get(0).ok_or(RuntimeError::Custom("Missing path argument".to_string()))?.as_string();
                 Ok(Value::String(std::path::Path::new(&path).parent().unwrap_or(std::path::Path::new("")).to_string_lossy().to_string()))
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
                self.side_effects.push_back(crate::RuntimeSideEffect::ConsoleOutput(final_msg));
                Ok(Value::Nothing)
            }
            "readline" => {
                // Console.ReadLine stub: no stdin, return empty string
                Ok(crate::builtins::console_fns::console_readline_fn())
            }
            "clear" => {
                self.side_effects.push_back(crate::RuntimeSideEffect::ConsoleClear);
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
            "floor" | "truncate" => crate::builtins::math_fns::floor_fn(args),
            "ceiling" => crate::builtins::math_fns::ceiling_fn(args),
            "pow" => crate::builtins::math_fns::pow_fn(args),
            "exp" => crate::builtins::math_fns::exp_fn(args),
            "log" => crate::builtins::math_fns::log_fn(args),
            "sin" => crate::builtins::math_fns::sin_fn(args),
            "cos" => crate::builtins::math_fns::cos_fn(args),
            "tan" => crate::builtins::math_fns::tan_fn(args),
            "asin" | "atan" | "atn" => crate::builtins::math_fns::atn_fn(args),
            "atan2" => crate::builtins::math_fns::atan2_fn(args),
            "sign" | "sgn" => crate::builtins::math_fns::sgn_fn(args),
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
