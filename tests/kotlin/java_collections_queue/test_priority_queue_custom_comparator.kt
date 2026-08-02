// vybe-test: kotlin/java_collections_queue/test_priority_queue_custom_comparator
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pq = java.util.PriorityQueue<String>(compareByDescending { it.length })
            pq.add("aa")
            pq.add("b")
            pq.add("ccc")
            __check((pq.poll()).toString(), "ccc")
            __check((pq.poll()).toString(), "aa")
            __check((pq.peek()).toString(), "b")
        }
