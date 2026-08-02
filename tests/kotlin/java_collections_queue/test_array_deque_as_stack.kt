// vybe-test: kotlin/java_collections_queue/test_array_deque_as_stack
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val stack = java.util.ArrayDeque<String>()
            stack.push("a")
            stack.push("b")
            __check((stack.pop()).toString(), "b")
            __check((stack.pop()).toString(), "a")
            __check((stack.isEmpty()).toString(), "true")
        }
