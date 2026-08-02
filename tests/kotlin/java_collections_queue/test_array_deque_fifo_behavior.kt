// vybe-test: kotlin/java_collections_queue/test_array_deque_fifo_behavior
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val q: java.util.ArrayDeque<Int> = java.util.ArrayDeque()
            q.addLast(1)
            q.addLast(2)
            q.addLast(3)
            __check((q.removeFirst()).toString(), "1")
            __check((q.peekFirst()).toString(), "2")
            __check((q.removeLast()).toString(), "3")
            __check((q.size).toString(), "1")
        }
