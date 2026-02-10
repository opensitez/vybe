# Collections Demo

This example demonstrates the new System.Collections.Generic types:

## Features Demonstrated

### Queue (FIFO - First In First Out)
- Task queue processing system
- `Enqueue()` - Add items to the back
- `Dequeue()` - Remove items from the front
- `Peek()` - Look at front item without removing
- `Count` - Number of items in queue
- `Clear()` - Remove all items

### Stack (LIFO - Last In First Out)
- Undo/Redo functionality
- `Push()` - Add items to the top
- `Pop()` - Remove items from the top
- `Peek()` - Look at top item without removing
- `Count` - Number of items in stack
- `Clear()` - Remove all items

### HashSet (Unique Items)
- Website visitor tracking
- `Add()` - Add unique items (returns True if new, False if duplicate)
- `Remove()` - Remove items
- `Contains()` - Check if item exists
- `Count` - Number of unique items
- `Clear()` - Remove all items

### Dictionary (Key-Value Pairs)
- User profile database
- `Add(key, value)` - Add key-value pairs
- `Item(key)` - Get value by key
- `ContainsKey(key)` - Check if key exists
- `Remove(key)` - Remove key-value pair
- `Count` - Number of key-value pairs
- `Clear()` - Remove all items

### Convert.ToDateTime
- Parse dates from strings
- Convert numeric date values
- Convert between date formats

## Running the Demo

```bash
cargo run --bin irys_editor -- run examples/collections_demo/CollectionsDemo.vbproj
```

Or from the test suite:
```bash
cargo run --bin irys_editor -- run tests/test_generic_collections.vb
```

## Expected Output

The demo shows practical use cases for each collection type with real-world scenarios like task queues, undo stacks, visitor tracking, and user databases.
