// vybe-test: kotlin/local_classes/test_local_class_list_composition
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Node(val value: Int)
            val nodes = listOf(Node(1), Node(2))
            __check((nodes[1].value).toString(), "2")
        }
