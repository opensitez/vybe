// vybe-test: kotlin/class_delegation/test_list_interface_delegation_size_and_index
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

class ReadOnlyList(delegate: List<Int>) : List<Int> by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val l = ReadOnlyList(listOf(2, 4, 6))
            __check((l.size).toString(), "3")
            __check((l[1]).toString(), "4")
            __check((l.contains(6)).toString(), "true")
        }
