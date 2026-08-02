// vybe-test: kotlin/scoping_functions/test_apply_keeps_reference_and_allows_multiple_mutations
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Box(var value: Int = 0)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Box()
                .apply { value += 1 }
                .apply { value += 2 }
                .apply { value *= 2 }
            __check((box.value).toString(), "6")
        }
