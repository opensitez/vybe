use crate::helpers::run_prints;

#[test]
fn test_array_deque_fifo_behavior() {
    let out = run_prints(r#"
        fun main() {
            val q: java.util.ArrayDeque<Int> = java.util.ArrayDeque()
            q.addLast(1)
            q.addLast(2)
            q.addLast(3)
            println(q.removeFirst())
            println(q.peekFirst())
            println(q.removeLast())
            println(q.size)
        }
    "#);
    assert_eq!(out, &["1", "2", "3", "1"]);
}

#[test]
fn test_array_deque_as_stack() {
    let out = run_prints(r#"
        fun main() {
            val stack = java.util.ArrayDeque<String>()
            stack.push("a")
            stack.push("b")
            println(stack.pop())
            println(stack.pop())
            println(stack.isEmpty())
        }
    "#);
    assert_eq!(out, &["b", "a", "true"]);
}

#[test]
fn test_array_deque_poll_offer_peek() {
    let out = run_prints(r#"
        fun main() {
            val q = java.util.ArrayDeque<Int>()
            q.offer(10)
            q.offer(20)
            println(q.peek())
            println(q.poll())
            println(q.poll())
            println(q.peek() ?: "none")
        }
    "#);
    assert_eq!(out, &["10", "10", "20", "none"]);
}

#[test]
fn test_priority_queue_orders_by_natural_order() {
    let out = run_prints(r#"
        fun main() {
            val pq = java.util.PriorityQueue<Int>()
            pq.add(5)
            pq.add(1)
            pq.add(3)
            println(pq.peek())
            println(pq.poll())
            println(pq.peek())
            println(pq.size)
        }
    "#);
    assert_eq!(out, &["1", "1", "3", "2"]);
}

#[test]
fn test_priority_queue_custom_comparator() {
    let out = run_prints(r#"
        fun main() {
            val pq = java.util.PriorityQueue<String>(compareByDescending { it.length })
            pq.add("aa")
            pq.add("b")
            pq.add("ccc")
            println(pq.poll())
            println(pq.poll())
            println(pq.peek())
        }
    "#);
    assert_eq!(out, &["ccc", "aa", "b"]);
}

#[test]
fn test_linked_list_as_queue() {
    let out = run_prints(r#"
        fun main() {
            val list = java.util.LinkedList<String>()
            list.offer("first")
            list.offer("second")
            println(list.peek())
            println(list.poll())
            println(list.peek())
            println(list.removeFirst())
        }
    "#);
    assert_eq!(out, &["first", "first", "second", "second"]);
}

#[test]
fn test_vector_iterator_and_set() {
    let out = run_prints(r#"
        fun main() {
            val v = java.util.Vector<Int>()
            v.add(1)
            v.add(2)
            v.addElement(3)
            val it = v.iterator()
            val sum = v.toMutableList().sum()
            println(v.elementAt(1))
            println(sum)
            println(it.next())
            println(v.size)
        }
    "#);
    assert_eq!(out, &["2", "6", "1", "3"]);
}

#[test]
fn test_stack_legacy_api_behavior() {
    let out = run_prints(r#"
        fun main() {
            val stack = java.util.Stack<Int>()
            stack.push(1)
            stack.push(2)
            println(stack.peek())
            println(stack.pop())
            println(stack.peek())
            println(stack.empty())
            println(stack.size)
        }
    "#);
    assert_eq!(out, &["2", "2", "1", "false", "1"]);
}

#[test]
fn test_arrays_deque_bidirectional_views() {
    let out = run_prints(r#"
        fun main() {
            val q = java.util.ArrayDeque<Int>()
            q.addFirst(1)
            q.addLast(2)
            q.addFirst(0)
            println(q.toString())
            println(q.pop())
            println(q.removeLast())
            println(q.removeFirst())
        }
    "#);
    assert_eq!(out, &["[0, 1, 2]", "0", "2", "1"]);
}
