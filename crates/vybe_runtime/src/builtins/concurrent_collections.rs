use std::sync::{Arc, RwLock, Mutex};
use std::collections::{HashMap, VecDeque};
use crate::value::{Value, SharedValue};

// ---------------------------------------------------------------------------
// ConcurrentDictionary - Thread-safe dictionary
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct ConcurrentDictionary {
    inner: Arc<RwLock<HashMap<String, SharedValue>>>,
}

impl ConcurrentDictionary {
    pub fn new() -> Self {
        Self {
            inner: Arc::new(RwLock::new(HashMap::new())),
        }
    }

    pub fn add_or_update(&self, key: &str, add_value: Value, update_value: Value) -> Value {
        let mut map = self.inner.write().unwrap();
        if map.contains_key(key) {
            // Update
            let shared = update_value.to_shared();
            map.insert(key.to_string(), shared);
            update_value
        } else {
            // Add
            let shared = add_value.to_shared();
            map.insert(key.to_string(), shared);
            add_value
        }
    }

    pub fn try_add(&self, key: &str, value: Value) -> bool {
        let mut map = self.inner.write().unwrap();
        if map.contains_key(key) {
            false
        } else {
            map.insert(key.to_string(), value.to_shared());
            true
        }
    }

    pub fn try_get_value(&self, key: &str) -> Option<Value> {
        let map = self.inner.read().unwrap();
        map.get(key).map(|v| v.to_value())
    }

    pub fn try_remove(&self, key: &str) -> Option<Value> {
        let mut map = self.inner.write().unwrap();
        map.remove(key).map(|v| v.to_value())
    }

    pub fn get_or_add(&self, key: &str, value: Value) -> Value {
        let mut map = self.inner.write().unwrap();
        if let Some(v) = map.get(key) {
            v.to_value()
        } else {
            let shared = value.to_shared();
            map.insert(key.to_string(), shared.clone());
            value
        }
    }

    pub fn count(&self) -> i32 {
        let map = self.inner.read().unwrap();
        map.len() as i32
    }

    pub fn clear(&self) {
        let mut map = self.inner.write().unwrap();
        map.clear();
    }

    pub fn contains_key(&self, key: &str) -> bool {
        let map = self.inner.read().unwrap();
        map.contains_key(key)
    }

    pub fn to_array(&self) -> Vec<Value> {
        let map = self.inner.read().unwrap();
        let mut items = Vec::new();
        for (k, v) in map.iter() {
            // Create KeyValuePair object
            let mut fields = HashMap::new();
            fields.insert("key".to_string(), Value::String(k.clone()));
            fields.insert("value".to_string(), v.to_value());
            fields.insert("__type".to_string(), Value::String("KeyValuePair".to_string()));
            
            let obj = crate::value::ObjectData { drawing_commands: Vec::new(),
                class_name: "KeyValuePair".to_string(),
                fields,
            };
            items.push(Value::Object(std::rc::Rc::new(std::cell::RefCell::new(obj))));
        }
        items
    }
    
    pub fn keys(&self) -> Vec<String> {
        let map = self.inner.read().unwrap();
        map.keys().cloned().collect()
    }

    pub fn values(&self) -> Vec<Value> {
        let map = self.inner.read().unwrap();
        map.values().map(|v| v.to_value()).collect()
    }
}

impl PartialEq for ConcurrentDictionary {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.inner, &other.inner)
    }
}

// ---------------------------------------------------------------------------
// ConcurrentQueue - Thread-safe FIFO queue
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct ConcurrentQueue {
    inner: Arc<Mutex<VecDeque<SharedValue>>>,
}

impl ConcurrentQueue {
    pub fn new() -> Self {
        Self {
            inner: Arc::new(Mutex::new(VecDeque::new())),
        }
    }

    pub fn enqueue(&self, value: Value) {
        let mut q = self.inner.lock().unwrap();
        q.push_back(value.to_shared());
    }

    pub fn try_dequeue(&self) -> Option<Value> {
        let mut q = self.inner.lock().unwrap();
        q.pop_front().map(|v| v.to_value())
    }

    pub fn try_peek(&self) -> Option<Value> {
        let q = self.inner.lock().unwrap();
        q.front().map(|v| v.to_value())
    }

    pub fn count(&self) -> i32 {
        let q = self.inner.lock().unwrap();
        q.len() as i32
    }

    pub fn clear(&self) {
        let mut q = self.inner.lock().unwrap();
        q.clear();
    }

    pub fn to_array(&self) -> Vec<Value> {
        let q = self.inner.lock().unwrap();
        q.iter().map(|v| v.to_value()).collect()
    }
}

impl PartialEq for ConcurrentQueue {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.inner, &other.inner)
    }
}

// ---------------------------------------------------------------------------
// ConcurrentStack - Thread-safe LIFO stack
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct ConcurrentStack {
    inner: Arc<Mutex<Vec<SharedValue>>>,
}

impl ConcurrentStack {
    pub fn new() -> Self {
        Self {
            inner: Arc::new(Mutex::new(Vec::new())),
        }
    }

    pub fn push(&self, value: Value) {
        let mut s = self.inner.lock().unwrap();
        s.push(value.to_shared());
    }

    pub fn try_pop(&self) -> Option<Value> {
        let mut s = self.inner.lock().unwrap();
        s.pop().map(|v| v.to_value())
    }

    pub fn try_peek(&self) -> Option<Value> {
        let s = self.inner.lock().unwrap();
        s.last().map(|v| v.to_value())
    }

    pub fn count(&self) -> i32 {
        let s = self.inner.lock().unwrap();
        s.len() as i32
    }

    pub fn clear(&self) {
        let mut s = self.inner.lock().unwrap();
        s.clear();
    }

    pub fn to_array(&self) -> Vec<Value> {
        let s = self.inner.lock().unwrap();
        s.iter().rev().map(|v| v.to_value()).collect()
    }
}

impl PartialEq for ConcurrentStack {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.inner, &other.inner)
    }
}
