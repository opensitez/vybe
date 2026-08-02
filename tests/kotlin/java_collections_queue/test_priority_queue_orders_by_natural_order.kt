// vybe-test: kotlin/java_collections_queue/test_priority_queue_orders_by_natural_order
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pq = java.util.PriorityQueue<Int>()
            pq.add(5)
            pq.add(1)
            pq.add(3)
            __check((pq.peek()).toString(), "1")
            __check((pq.poll()).toString(), "1")
            __check((pq.peek()).toString(), "3")
            __check((pq.size).toString(), "2")
        }
