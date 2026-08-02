// vybe-test: kotlin/java_collections_queue/test_array_deque_poll_offer_peek
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val q = java.util.ArrayDeque<Int>()
            q.offer(10)
            q.offer(20)
            __check((q.peek()).toString(), "10")
            __check((q.poll()).toString(), "10")
            __check((q.poll()).toString(), "20")
            __check((q.peek() ?: "none").toString(), "none")
        }
