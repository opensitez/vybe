use crate::value::{Value, RuntimeError};
use std::collections::VecDeque;

// ---------------------------------------------------------------------------
// ArrayList  (already existed)
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Queue  — FIFO collection
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
pub struct Queue {
    items: VecDeque<Value>,
}

impl Queue {
    pub fn new() -> Self {
        Self { items: VecDeque::new() }
    }

    /// Add to the back.
    pub fn enqueue(&mut self, value: Value) {
        self.items.push_back(value);
    }

    /// Remove and return the front item.
    pub fn dequeue(&mut self) -> Result<Value, RuntimeError> {
        self.items.pop_front().ok_or_else(|| RuntimeError::Custom("Queue is empty".to_string()))
    }

    /// Return the front item without removing it.
    pub fn peek(&self) -> Result<Value, RuntimeError> {
        self.items.front().cloned().ok_or_else(|| RuntimeError::Custom("Queue is empty".to_string()))
    }

    pub fn count(&self) -> i32 {
        self.items.len() as i32
    }

    pub fn clear(&mut self) {
        self.items.clear();
    }

    pub fn contains(&self, value: &Value) -> bool {
        self.items.contains(value)
    }

    pub fn to_array(&self) -> Vec<Value> {
        self.items.iter().cloned().collect()
    }
}

// ---------------------------------------------------------------------------
// Stack  — LIFO collection
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
pub struct Stack {
    items: Vec<Value>,
}

impl Stack {
    pub fn new() -> Self {
        Self { items: Vec::new() }
    }

    /// Push onto the top.
    pub fn push(&mut self, value: Value) {
        self.items.push(value);
    }

    /// Remove and return the top item.
    pub fn pop(&mut self) -> Result<Value, RuntimeError> {
        self.items.pop().ok_or_else(|| RuntimeError::Custom("Stack is empty".to_string()))
    }

    /// Return the top item without removing it.
    pub fn peek(&self) -> Result<Value, RuntimeError> {
        self.items.last().cloned().ok_or_else(|| RuntimeError::Custom("Stack is empty".to_string()))
    }

    pub fn count(&self) -> i32 {
        self.items.len() as i32
    }

    pub fn clear(&mut self) {
        self.items.clear();
    }

    pub fn contains(&self, value: &Value) -> bool {
        self.items.contains(value)
    }

    pub fn to_array(&self) -> Vec<Value> {
        // Stack.ToArray returns top-first order
        self.items.iter().rev().cloned().collect()
    }
}

// ---------------------------------------------------------------------------
// HashSet  — unique-value collection
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
pub struct VBHashSet {
    items: Vec<Value>, // Vec to preserve insertion order; uniqueness checked manually
}

impl VBHashSet {
    pub fn new() -> Self {
        Self { items: Vec::new() }
    }

    /// Add a value. Returns true if the value was new, false if it already existed.
    pub fn add(&mut self, value: Value) -> bool {
        if self.items.contains(&value) {
            false
        } else {
            self.items.push(value);
            true
        }
    }

    pub fn remove(&mut self, value: &Value) -> bool {
        if let Some(pos) = self.items.iter().position(|x| x == value) {
            self.items.remove(pos);
            true
        } else {
            false
        }
    }

    pub fn contains(&self, value: &Value) -> bool {
        self.items.contains(value)
    }

    pub fn count(&self) -> i32 {
        self.items.len() as i32
    }

    pub fn clear(&mut self) {
        self.items.clear();
    }

    pub fn to_array(&self) -> Vec<Value> {
        self.items.clone()
    }
}

// ---------------------------------------------------------------------------
// Dictionary  — key/value collection
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
pub struct VBDictionary {
    // Parallel vecs to preserve insertion order and allow any Value as key.
    keys: Vec<Value>,
    values: Vec<Value>,
}

impl VBDictionary {
    pub fn new() -> Self {
        Self { keys: Vec::new(), values: Vec::new() }
    }

    fn find_key(&self, key: &Value) -> Option<usize> {
        // Case-insensitive string comparison for string keys, otherwise equality
        self.keys.iter().position(|k| Self::keys_equal(k, key))
    }

    fn keys_equal(a: &Value, b: &Value) -> bool {
        match (a, b) {
            (Value::String(sa), Value::String(sb)) => sa.eq_ignore_ascii_case(sb),
            _ => a == b,
        }
    }

    pub fn add(&mut self, key: Value, value: Value) -> Result<(), RuntimeError> {
        if self.find_key(&key).is_some() {
            Err(RuntimeError::Custom(format!("An item with the same key has already been added: {}", key.as_string())))
        } else {
            self.keys.push(key);
            self.values.push(value);
            Ok(())
        }
    }

    pub fn item(&self, key: &Value) -> Result<Value, RuntimeError> {
        if let Some(idx) = self.find_key(key) {
            Ok(self.values[idx].clone())
        } else {
            Err(RuntimeError::Custom(format!("The given key was not present in the dictionary: {}", key.as_string())))
        }
    }

    pub fn set_item(&mut self, key: Value, value: Value) {
        if let Some(idx) = self.find_key(&key) {
            self.values[idx] = value;
        } else {
            self.keys.push(key);
            self.values.push(value);
        }
    }

    pub fn contains_key(&self, key: &Value) -> bool {
        self.find_key(key).is_some()
    }

    pub fn contains_value(&self, value: &Value) -> bool {
        self.values.contains(value)
    }

    pub fn remove(&mut self, key: &Value) -> bool {
        if let Some(idx) = self.find_key(key) {
            self.keys.remove(idx);
            self.values.remove(idx);
            true
        } else {
            false
        }
    }

    pub fn count(&self) -> i32 {
        self.keys.len() as i32
    }

    pub fn clear(&mut self) {
        self.keys.clear();
        self.values.clear();
    }

    pub fn keys(&self) -> Vec<Value> {
        self.keys.clone()
    }

    pub fn values(&self) -> Vec<Value> {
        self.values.clone()
    }
}
