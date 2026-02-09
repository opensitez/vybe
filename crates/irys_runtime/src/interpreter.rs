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
                    _ => return Err(RuntimeError::Custom("For Each requires an array or string".to_string())),
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
            "chrw" => return chr_fn(&arg_values),
            "asc" | "ascw" => return asc_fn(&arg_values),
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
            "formatnumber" => return formatnumber_fn(&arg_values),
            "formatcurrency" => return formatcurrency_fn(&arg_values),
            "formatpercent" => return formatpercent_fn(&arg_values),

            // Misc utility functions
            "doevents" => return doevents_fn(&arg_values),
            "rgb" => return rgb_fn(&arg_values),
            "qbcolor" => return qbcolor_fn(&arg_values),
            "environ" | "environ$" => return environ_fn(&arg_values),
            "isnullorempty" | "string.isnullorempty" => return isnullorempty_fn(&arg_values),
            "dateadd" => return dateadd_fn(&arg_values),
            "datediff" => return datediff_fn(&arg_values),
            "strdup" => return strdup_fn(&arg_values),

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

        // Evaluate object to check if it's a Collection
        if let Ok(obj_val) = self.evaluate_expr(obj) {
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
                    eprintln!("DEBUG: Param {} is_byref: {}", param.name.as_str(), is_byref);
                    
                    if is_byref {
                        // Check if argument is a variable we can write back to
                        match &exprs[i] {
                            Expression::Variable(name) => {
                                eprintln!("DEBUG: Enabling writeback for {} -> {}", name.as_str(), param.name.as_str());
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
                     eprintln!("DEBUG: Capturing writeback value for {}: {:?}", caller_var, new_val);
                     final_writebacks.push((caller_var, new_val));
                 }
             }
        }

        // Pop scope
        self.env.pop_scope();
        self.current_object = prev_object;
        
        // Apply writebacks in restored scope
        for (var_name, val) in final_writebacks {
            eprintln!("DEBUG: Writing back {} = {:?}", var_name, val);
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
