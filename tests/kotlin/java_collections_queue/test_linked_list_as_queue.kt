// vybe-test: kotlin/java_collections_queue/test_linked_list_as_queue
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = java.util.LinkedList<String>()
            list.offer("first")
            list.offer("second")
            __check((list.peek()).toString(), "first")
            __check((list.poll()).toString(), "first")
            __check((list.peek()).toString(), "second")
            __check((list.removeFirst()).toString(), "second")
        }
