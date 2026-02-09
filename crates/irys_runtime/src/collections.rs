use crate::value::{Value, RuntimeError};
// use std::rc::Rc;
// use std::cell::RefCell;

#[derive(Debug, Clone, PartialEq)]
pub struct ArrayList {
    pub items: Vec<Value>,
}

impl ArrayList {
    pub fn new() -> Self {
        Self { items: Vec::new() }
    }

    pub fn add(&mut self, value: Value) -> i32 {
        self.items.push(value);
        (self.items.len() - 1) as i32
    }

    pub fn remove(&mut self, value: &Value) {
        if let Some(pos) = self.items.iter().position(|x| x == value) {
            self.items.remove(pos);
        }
    }

    pub fn remove_at(&mut self, index: usize) -> Result<(), RuntimeError> {
        if index < self.items.len() {
            self.items.remove(index);
            Ok(())
        } else {
            Err(RuntimeError::Custom(format!("Index out of range: {}", index)))
        }
    }

    pub fn clear(&mut self) {
        self.items.clear();
    }

    pub fn count(&self) -> i32 {
        self.items.len() as i32
    }

    pub fn item(&self, index: usize) -> Result<Value, RuntimeError> {
        self.items.get(index).cloned().ok_or_else(|| RuntimeError::Custom(format!("Index out of range: {}", index)))
    }
    
    pub fn set_item(&mut self, index: usize, value: Value) -> Result<(), RuntimeError> {
         if index < self.items.len() {
             self.items[index] = value;
             Ok(())
         } else {
             Err(RuntimeError::Custom(format!("Index out of range: {}", index)))
         }
    }
}
