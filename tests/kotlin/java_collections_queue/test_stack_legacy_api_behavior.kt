// vybe-test: kotlin/java_collections_queue/test_stack_legacy_api_behavior
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val stack = java.util.Stack<Int>()
            stack.push(1)
            stack.push(2)
            __check((stack.peek()).toString(), "2")
            __check((stack.pop()).toString(), "2")
            __check((stack.peek()).toString(), "1")
            __check((stack.empty()).toString(), "false")
            __check((stack.size).toString(), "1")
        }
