use crate::value::{RuntimeError, Value};
use std::collections::{HashMap, HashSet};

#[derive(Debug, Clone, PartialEq)]
pub struct Environment {
    scopes: Vec<HashMap<String, Value>>,
    constants: HashSet<String>, // Track constant names across all scopes
}

impl Environment {
    pub fn new() -> Self {
        Self {
            scopes: vec![HashMap::new()],
            constants: HashSet::new(),
        }
    }

    pub fn push_scope(&mut self) {
        self.scopes.push(HashMap::new());
    }

    pub fn pop_scope(&mut self) {
        if self.scopes.len() > 1 {
            self.scopes.pop();
        }
    }

    pub fn define(&mut self, name: impl Into<String>, value: Value) {
        let name = name.into().to_lowercase();
        if let Some(scope) = self.scopes.last_mut() {
            scope.insert(name, value);
        }
    }

    pub fn define_global(&mut self, name: impl Into<String>, value: Value) {
        let name = name.into().to_lowercase();
        if let Some(scope) = self.scopes.first_mut() {
            scope.insert(name, value);
        }
    }

    pub fn define_const(&mut self, name: impl Into<String>, value: Value) {
        let name = name.into().to_lowercase();
        self.constants.insert(name.clone());
        if let Some(scope) = self.scopes.last_mut() {
            scope.insert(name, value);
        }
    }

    pub fn set(&mut self, name: &str, value: Value) -> Result<(), RuntimeError> {
        let name_lower = name.to_lowercase();

        // Check if this is a constant
        if self.constants.contains(&name_lower) {
            return Err(RuntimeError::Custom(format!(
                "Cannot assign to constant '{}'",
                name
            )));
        }

        // Search from innermost to outermost scope
        for scope in self.scopes.iter_mut().rev() {
            if scope.contains_key(&name_lower) {
                scope.insert(name_lower, value);
                return Ok(());
            }
        }

        // If not found, define in current scope
        self.define(name_lower, value);
        Ok(())
    }

    pub fn get(&self, name: &str) -> Result<Value, RuntimeError> {
        let name_lower = name.to_lowercase();
        // Search from innermost to outermost scope
        for scope in self.scopes.iter().rev() {
            if let Some(value) = scope.get(&name_lower) {
                return Ok(value.clone());
            }
        }

        Err(RuntimeError::UndefinedVariable(name.to_string()))
    }

    pub fn get_or_nothing(&self, name: &str) -> Value {
        self.get(name).unwrap_or(Value::Nothing)
    }

    pub fn exists_in_current_scope(&self, name: &str) -> bool {
        let name_lower = name.to_lowercase();
        if let Some(scope) = self.scopes.last() {
            scope.contains_key(&name_lower)
        } else {
            false
        }
    }

    /// Check if a variable exists in any scope EXCEPT the global scope (index 0)
    pub fn has_local(&self, name: &str) -> bool {
        let name_lower = name.to_lowercase();
        if self.scopes.len() <= 1 {
            return false;
        }
        for scope in self.scopes.iter().skip(1) {
            if scope.contains_key(&name_lower) {
                return true;
            }
        }
        false
    }

    /// Get a variable from the global scope ONLY
    pub fn get_global(&self, name: &str) -> Option<Value> {
        let name_lower = name.to_lowercase();
        self.scopes.first().and_then(|s| s.get(&name_lower)).cloned()
    }

    /// Deep clone the environment for snapshot threading.
    pub fn deep_clone(&self) -> Self {
        let new_scopes = self.scopes.iter().map(|scope| {
            let mut new_scope = HashMap::new();
            for (k, v) in scope {
                new_scope.insert(k.clone(), v.deep_clone());
            }
            new_scope
        }).collect();
        
        Self {
            scopes: new_scopes,
            constants: self.constants.clone(),
        }
    }

    pub fn to_shared(&self) -> crate::value::SharedEnvironment {
        let new_scopes = self.scopes.iter().map(|scope| {
            let mut new_scope = HashMap::new();
            for (k, v) in scope {
                new_scope.insert(k.clone(), v.to_shared());
            }
            new_scope
        }).collect();
        
        crate::value::SharedEnvironment {
            scopes: new_scopes,
            constants: self.constants.clone(),
        }
    }
}

impl crate::value::SharedEnvironment {
    pub fn to_environment(&self) -> Environment {
        let new_scopes = self.scopes.iter().map(|scope| {
            let mut new_scope = HashMap::new();
            for (k, v) in scope {
                new_scope.insert(k.clone(), v.to_value());
            }
            new_scope
        }).collect();
        
        Environment {
            scopes: new_scopes,
            constants: self.constants.clone(),
        }
    }
}

impl Default for Environment {
    fn default() -> Self {
        Self::new()
    }
}
