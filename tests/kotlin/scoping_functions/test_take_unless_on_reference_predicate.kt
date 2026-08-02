// vybe-test: kotlin/scoping_functions/test_take_unless_on_reference_predicate
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Box(var n: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Box(11)
            val filtered = value.takeUnless { it.n % 2 == 1 }
            __check((filtered == null).toString(), "true")
            __check((value.n).toString(), "11")
            val keep = Box(4).takeUnless { it.n % 2 == 1 }
            __check((keep?.n).toString(), "4")
        }
