// vybe-test: kotlin/repeat_statements/test_repeat_nested_with_break_condition
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var out = 0
        repeat(5) { outer ->
            repeat(3) { inner ->
                if (outer == 2 && inner == 2) return@repeat
                out += 1
            }
        }
        __check((out).toString(), "14")
    }
