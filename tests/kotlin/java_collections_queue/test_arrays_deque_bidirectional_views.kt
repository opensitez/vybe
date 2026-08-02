// vybe-test: kotlin/java_collections_queue/test_arrays_deque_bidirectional_views
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val q = java.util.ArrayDeque<Int>()
            q.addFirst(1)
            q.addLast(2)
            q.addFirst(0)
            __check((q.toString()).toString(), "[0, 1, 2]")
            __check((q.pop()).toString(), "0")
            __check((q.removeLast()).toString(), "2")
            __check((q.removeFirst()).toString(), "1")
        }
