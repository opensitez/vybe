use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: heapq + bisect + array — heap operations, binary search, typed arrays
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_heapq_basic_heap_operations() {
    let src = r#"
import heapq

heap = []
for val in [5, 1, 8, 3, 9, 2]:
    heapq.heappush(heap, val)

print(heap[0])   # smallest element always at index 0
results = []
while heap:
    results.append(heapq.heappop(heap))
print(results)
"#;
    assert_eq!(run_python(src), vec!["1", "[1, 2, 3, 5, 8, 9]"]);
}

#[test]
fn test_py_heapq_heapify() {
    let src = r#"
import heapq

data = [5, 1, 8, 3, 9, 2, 7, 4, 6]
heapq.heapify(data)
print(data[0])  # min at root
print(heapq.heappop(data))
print(heapq.heappop(data))
"#;
    assert_eq!(run_python(src), vec!["1", "1", "2"]);
}

#[test]
fn test_py_heapq_nlargest_nsmallest() {
    let src = r#"
import heapq

data = [3, 1, 4, 1, 5, 9, 2, 6, 5, 3, 5]
print(heapq.nlargest(3, data))
print(heapq.nsmallest(3, data))
"#;
    assert_eq!(run_python(src), vec!["[9, 6, 5]", "[1, 1, 2]"]);
}

#[test]
fn test_py_heapq_with_priority_tuples() {
    let src = r#"
import heapq

# Priority queue using (priority, item) tuples
pq = []
heapq.heappush(pq, (3, "low"))
heapq.heappush(pq, (1, "high"))
heapq.heappush(pq, (2, "medium"))

results = []
while pq:
    priority, item = heapq.heappop(pq)
    results.append(item)
print(results)
"#;
    assert_eq!(run_python(src), vec!["['high', 'medium', 'low']"]);
}

#[test]
fn test_py_heapq_merge_sorted() {
    let src = r#"
import heapq

a = [1, 4, 7]
b = [2, 5, 8]
c = [3, 6, 9]
merged = list(heapq.merge(a, b, c))
print(merged)
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4, 5, 6, 7, 8, 9]"]);
}

#[test]
fn test_py_heapq_nlargest_with_key() {
    let src = r#"
import heapq

students = [
    {"name": "Alice", "gpa": 3.8},
    {"name": "Bob", "gpa": 3.5},
    {"name": "Charlie", "gpa": 3.9},
    {"name": "Dave", "gpa": 3.2},
]

top2 = heapq.nlargest(2, students, key=lambda s: s["gpa"])
print([s["name"] for s in top2])
"#;
    assert_eq!(run_python(src), vec!["['Charlie', 'Alice']"]);
}

#[test]
fn test_py_bisect_search_insert_point() {
    let src = r#"
import bisect

data = [1, 3, 5, 7, 9, 11]
print(bisect.bisect_left(data, 5))   # index of first 5
print(bisect.bisect_right(data, 5))  # index after last 5
print(bisect.bisect_left(data, 6))   # where 6 would go
print(bisect.bisect_right(data, 4))
"#;
    assert_eq!(run_python(src), vec!["2", "3", "3", "2"]);
}

#[test]
fn test_py_bisect_insort() {
    let src = r#"
import bisect

data = [1, 3, 5, 7, 9]
bisect.insort(data, 4)
bisect.insort(data, 6)
bisect.insort(data, 0)
print(data)
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 3, 4, 5, 6, 7, 9]"]);
}

#[test]
fn test_py_bisect_grade_lookup() {
    let src = r#"
import bisect

grades = [("F", 60), ("D", 65), ("C", 70), ("B", 80), ("A", 90)]
breakpoints = [bp for _, bp in grades]
letters = [letter for letter, _ in grades]

def grade(score):
    idx = bisect.bisect_left(breakpoints, score)
    if idx >= len(letters):
        return "A+"
    return letters[idx]

print(grade(55))
print(grade(65))
print(grade(80))
print(grade(95))
"#;
    assert_eq!(run_python(src), vec!["F", "D", "B", "A+"]);
}

#[test]
fn test_py_array_module_typed_array() {
    let src = r#"
import array

arr = array.array("i", [1, 2, 3, 4, 5])
print(arr.typecode)
print(arr[0])
arr.append(6)
print(len(arr))
print(arr.itemsize)

total = sum(arr)
print(total)
"#;
    assert_eq!(run_python(src), vec!["i", "1", "6", "4", "21"]);
}

#[test]
fn test_py_array_tobytes_frombytes() {
    let src = r#"
import array

original = array.array("d", [1.0, 2.0, 3.14])
data = original.tobytes()
print(len(data))  # 3 doubles * 8 bytes each

restored = array.array("d")
restored.frombytes(data)
print(list(restored[:2]))
print(round(restored[2], 2))
"#;
    assert_eq!(run_python(src), vec!["24", "[1.0, 2.0]", "3.14"]);
}

#[test]
fn test_py_heapq_pushpop_and_replace() {
    let src = r#"
import heapq

heap = [1, 3, 5, 7]
heapq.heapify(heap)

# pushpop: more efficient than push+pop
result = heapq.heappushpop(heap, 4)
print(result)   # returns min of (pushed, heap min)
print(heap[0])

# heapreplace: pop min then push
popped = heapq.heapreplace(heap, 0)
print(popped)
print(heap[0])
"#;
    assert_eq!(run_python(src), vec!["1", "3", "3", "0"]);
}
