// vybe-test: kotlin/object_expressions/test_object_expression_with_mutable_interface_property
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface Tally {
    fun next(): Int
    fun total(): Int
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val tally = object : Tally {
        var count = 0
        override fun next(): Int {
            count += 1
            return count
        }

        override fun total(): Int = count
    }

    __check((tally.next()).toString(), "1")
    __check((tally.next()).toString(), "2")
    __check((tally.total()).toString(), "2")
}
