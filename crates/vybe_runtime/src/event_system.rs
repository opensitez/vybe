use std::collections::HashMap;
use vybe_forms::EventType;

pub type EventHandler = Box<dyn Fn() + Send + Sync>;

#[derive(Default)]
pub struct EventSystem {
    handlers: HashMap<String, Vec<String>>, // (control_name, event_type) -> list of handler_names
}

impl EventSystem {
    pub fn new() -> Self {
        Self {
            handlers: HashMap::new(),
        }
    }

    pub fn register(&mut self, control_name: impl Into<String>, event_type: EventType, handler_name: impl Into<String>) {
        let key = format!("{}_{}", control_name.into().to_lowercase(), event_type.as_str().to_lowercase());
        self.handlers.entry(key).or_default().push(handler_name.into());
    }

    /// Alias used by AddHandler statement
    pub fn register_handler(&mut self, control_name: &str, event_type: &EventType, handler_name: &str) {
        let key = format!("{}_{}", control_name.to_lowercase(), event_type.as_str().to_lowercase());
        self.handlers.entry(key).or_default().push(handler_name.to_string());
    }

    /// Remove a handler (used by RemoveHandler statement)
    pub fn remove_handler(&mut self, control_name: &str, event_type: &EventType, handler_name: &str) {
        let key = format!("{}_{}", control_name.to_lowercase(), event_type.as_str().to_lowercase());
        if let Some(list) = self.handlers.get_mut(&key) {
            // Remove the first occurrence of the handler logic (VB.NET behavior)
            // Note: In .NET delegates are compared by target/method. Here we compare by name.
            if let Some(pos) = list.iter().position(|x| x.eq_ignore_ascii_case(handler_name)) {
                list.remove(pos);
            }
        }
    }

    pub fn get_handlers(&self, control_name: &str, event_type: &EventType) -> Option<&Vec<String>> {
        let key = format!("{}_{}", control_name.to_lowercase(), event_type.as_str().to_lowercase());
        self.handlers.get(&key)
    }
}
