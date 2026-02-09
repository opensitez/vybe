use std::collections::HashMap;
use irys_forms::EventType;

pub type EventHandler = Box<dyn Fn() + Send + Sync>;

#[derive(Default)]
pub struct EventSystem {
    handlers: HashMap<String, String>, // (control_name, event_type) -> handler_name
}

impl EventSystem {
    pub fn new() -> Self {
        Self {
            handlers: HashMap::new(),
        }
    }

    pub fn register(&mut self, control_name: impl Into<String>, event_type: EventType, handler_name: impl Into<String>) {
        let key = format!("{}_{}", control_name.into().to_lowercase(), event_type.as_str().to_lowercase());
        self.handlers.insert(key, handler_name.into());
    }

    pub fn get_handler(&self, control_name: &str, event_type: &EventType) -> Option<&str> {
        let key = format!("{}_{}", control_name.to_lowercase(), event_type.as_str().to_lowercase());
        self.handlers.get(&key).map(|s| s.as_str())
    }
}
