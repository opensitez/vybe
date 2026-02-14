use crate::value::{Value, RuntimeError};
use std::collections::VecDeque;
use std::collections::HashMap;

// ---------------------------------------------------------------------------
// ArrayList  — also serves as VB.NET Collection (with optional string keys)
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
pub struct ArrayList {
    pub items: Vec<Value>,
    /// Optional key→index mapping for VB.NET Collection-style keyed access.
    /// Keys are stored lower-cased for case-insensitive lookup.
    pub keys: HashMap<String, usize>,
}

impl ArrayList {
    pub fn new() -> Self {
        Self { items: Vec::new(), keys: HashMap::new() }
    }

    pub fn add(&mut self, value: Value) -> i32 {
        self.items.push(value);
        (self.items.len() - 1) as i32
    }

    /// VB.NET Collection.Add(item, key) — add with a string key
    pub fn add_with_key(&mut self, value: Value, key: &str) -> Result<i32, RuntimeError> {
        let key_lower = key.to_lowercase();
        if self.keys.contains_key(&key_lower) {
            return Err(RuntimeError::Custom(format!(
                "Argument 'Key' is not valid. Duplicate key: '{}'", key
            )));
        }
        let idx = self.items.len();
        self.items.push(value);
        self.keys.insert(key_lower, idx);
        Ok(idx as i32)
    }

    /// VB.NET Collection.Add(item, key, before, after) — full signature
    pub fn add_with_key_position(&mut self, value: Value, key: Option<&str>, before: Option<usize>, after: Option<usize>) -> Result<i32, RuntimeError> {
        let insert_pos = if let Some(b) = before {
            // VB.NET Collection is 1-based for before/after
            if b == 0 || b > self.items.len() + 1 {
                return Err(RuntimeError::Custom(format!("Argument 'Before' out of range: {}", b)));
            }
            b - 1
        } else if let Some(a) = after {
            if a == 0 || a > self.items.len() {
                return Err(RuntimeError::Custom(format!("Argument 'After' out of range: {}", a)));
            }
            a // after is 1-based, so insert at position = after
        } else {
            self.items.len()
        };

        if let Some(k) = key {
            let key_lower = k.to_lowercase();
            if self.keys.contains_key(&key_lower) {
                return Err(RuntimeError::Custom(format!(
                    "Argument 'Key' is not valid. Duplicate key: '{}'", k
                )));
            }
            // Shift existing key indices that are >= insert_pos
            for v in self.keys.values_mut() {
                if *v >= insert_pos {
                    *v += 1;
                }
            }
            self.keys.insert(key_lower, insert_pos);
        } else {
            // Still shift existing keys
            for v in self.keys.values_mut() {
                if *v >= insert_pos {
                    *v += 1;
                }
            }
        }

        self.items.insert(insert_pos, value);
        Ok(insert_pos as i32)
    }

    pub fn remove(&mut self, value: &Value) {
        if let Some(pos) = self.items.iter().position(|x| x == value) {
            self.items.remove(pos);
            self.rebuild_key_indices_after_remove(pos);
        }
    }

    /// VB.NET Collection.Remove(key) — remove by string key
    pub fn remove_by_key(&mut self, key: &str) -> Result<(), RuntimeError> {
        let key_lower = key.to_lowercase();
        if let Some(idx) = self.keys.remove(&key_lower) {
            if idx < self.items.len() {
                self.items.remove(idx);
                self.rebuild_key_indices_after_remove(idx);
            }
            Ok(())
        } else {
            Err(RuntimeError::Custom(format!("Argument 'Key' is not valid. Key not found: '{}'", key)))
        }
    }

    pub fn remove_at(&mut self, index: usize) -> Result<(), RuntimeError> {
        if index < self.items.len() {
            self.items.remove(index);
            self.rebuild_key_indices_after_remove(index);
            Ok(())
        } else {
            Err(RuntimeError::Custom(format!("Index out of range: {}", index)))
        }
    }

    /// After removing an item at `removed_pos`, fix up key indices.
    fn rebuild_key_indices_after_remove(&mut self, removed_pos: usize) {
        // Remove any key that pointed to removed_pos
        self.keys.retain(|_, v| *v != removed_pos);
        // Shift down indices above removed_pos
        for v in self.keys.values_mut() {
            if *v > removed_pos {
                *v -= 1;
            }
        }
    }

    pub fn clear(&mut self) {
        self.items.clear();
        self.keys.clear();
    }

    pub fn count(&self) -> i32 {
        self.items.len() as i32
    }

    pub fn item(&self, index: usize) -> Result<Value, RuntimeError> {
        self.items.get(index).cloned().ok_or_else(|| RuntimeError::Custom(format!("Index out of range: {}", index)))
    }

    /// VB.NET Collection.Item(key) — retrieve by string key
    pub fn item_by_key(&self, key: &str) -> Result<Value, RuntimeError> {
        let key_lower = key.to_lowercase();
        if let Some(&idx) = self.keys.get(&key_lower) {
            self.items.get(idx).cloned().ok_or_else(|| {
                RuntimeError::Custom(format!("Key index out of range: '{}' -> {}", key, idx))
            })
        } else {
            Err(RuntimeError::Custom(format!("Argument 'Index' is not valid. Key not found: '{}'", key)))
        }
    }

    /// Check if a string key exists.
    pub fn contains_key(&self, key: &str) -> bool {
        self.keys.contains_key(&key.to_lowercase())
    }

    pub fn set_item(&mut self, index: usize, value: Value) -> Result<(), RuntimeError> {
         if index < self.items.len() {
             self.items[index] = value;
             Ok(())
         } else {
             Err(RuntimeError::Custom(format!("Index out of range: {}", index)))
         }
    }

    /// Capacity property — returns current Vec capacity.
    pub fn capacity(&self) -> i32 {
        self.items.capacity() as i32
    }

    /// Set the capacity (reserve at least this many slots).
    pub fn set_capacity(&mut self, cap: usize) {
        if cap > self.items.len() {
            self.items.reserve(cap - self.items.len());
        }
    }

    /// TrimToSize — shrink capacity to match count.
    pub fn trim_to_size(&mut self) {
        self.items.shrink_to_fit();
    }

    /// Clone — shallow copy.
    pub fn clone_list(&self) -> Self {
        Self {
            items: self.items.clone(),
            keys: self.keys.clone(),
        }
    }

    /// GetRange(index, count) — return a sub-list.
    pub fn get_range(&self, index: usize, count: usize) -> Result<Vec<Value>, RuntimeError> {
        if index + count > self.items.len() {
            return Err(RuntimeError::Custom(format!(
                "GetRange: index {} + count {} exceeds length {}", index, count, self.items.len()
            )));
        }
        Ok(self.items[index..index + count].to_vec())
    }

    /// InsertRange(index, collection) — insert multiple items at index.
    pub fn insert_range(&mut self, index: usize, items: Vec<Value>) {
        let insert_at = index.min(self.items.len());
        // Shift key indices >= insert_at
        let shift = items.len();
        for v in self.keys.values_mut() {
            if *v >= insert_at {
                *v += shift;
            }
        }
        for (i, item) in items.into_iter().enumerate() {
            self.items.insert(insert_at + i, item);
        }
    }

    /// RemoveRange(index, count) — remove a range of elements.
    pub fn remove_range(&mut self, index: usize, count: usize) -> Result<(), RuntimeError> {
        if index + count > self.items.len() {
            return Err(RuntimeError::Custom(format!(
                "RemoveRange: index {} + count {} exceeds length {}", index, count, self.items.len()
            )));
        }
        self.items.drain(index..index + count);
        // Rebuild keys: remove any pointing into removed range, shift rest
        self.keys.retain(|_, v| *v < index || *v >= index + count);
        for v in self.keys.values_mut() {
            if *v >= index + count {
                *v -= count;
            }
        }
        Ok(())
    }

    /// SetRange(index, collection) — overwrite elements starting at index.
    pub fn set_range(&mut self, index: usize, items: &[Value]) -> Result<(), RuntimeError> {
        if index + items.len() > self.items.len() {
            return Err(RuntimeError::Custom(format!(
                "SetRange: index {} + count {} exceeds length {}", index, items.len(), self.items.len()
            )));
        }
        for (i, item) in items.iter().enumerate() {
            self.items[index + i] = item.clone();
        }
        Ok(())
    }

    /// Reverse a subrange.
    pub fn reverse_range(&mut self, index: usize, count: usize) -> Result<(), RuntimeError> {
        if index + count > self.items.len() {
            return Err(RuntimeError::Custom(format!(
                "Reverse: index {} + count {} exceeds length {}", index, count, self.items.len()
            )));
        }
        self.items[index..index + count].reverse();
        Ok(())
    }

    /// BinarySearch — assumes the list is sorted. Returns index or bitwise complement of insertion point.
    pub fn binary_search(&self, value: &Value) -> i32 {
        let result = self.items.binary_search_by(|item| {
            let a = item.as_string();
            let b = value.as_string();
            a.cmp(&b)
        });
        match result {
            Ok(idx) => idx as i32,
            Err(idx) => !(idx as i32), // bitwise complement like .NET
        }
    }

    /// CopyTo(destination_start_index) — copies items into an existing array.
    /// Returns the items as a Vec for the interpreter to place into the target array.
    pub fn copy_to(&self) -> Vec<Value> {
        self.items.clone()
    }

    /// IndexOf with start index.
    pub fn index_of_from(&self, value: &Value, start: usize) -> i32 {
        for i in start..self.items.len() {
            if self.items[i] == *value {
                return i as i32;
            }
        }
        -1
    }

    /// IndexOf with start index and count.
    pub fn index_of_range(&self, value: &Value, start: usize, count: usize) -> i32 {
        let end = (start + count).min(self.items.len());
        for i in start..end {
            if self.items[i] == *value {
                return i as i32;
            }
        }
        -1
    }

    /// LastIndexOf with start index.
    pub fn last_index_of_from(&self, value: &Value, start: usize) -> i32 {
        let end = start.min(self.items.len().saturating_sub(1));
        for i in (0..=end).rev() {
            if self.items[i] == *value {
                return i as i32;
            }
        }
        -1
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

    pub fn from_vecdeque(items: VecDeque<Value>) -> Self {
        Self { items }
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

    pub fn from_vec(items: Vec<Value>) -> Self {
        Self { items }
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

    pub fn from_vec(items: Vec<Value>) -> Self {
        Self { items }
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

    pub fn from_parts(keys: Vec<Value>, values: Vec<Value>) -> Self {
        Self { keys, values }
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
