// vybe-test: kotlin/math_builtins/test_math_pipeline_with_floor_and_round
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val raw = abs(floor(-3.7) + ceil(3.2))
            val rounded = round(raw / 2.0)
            __check((raw).toString(), "7")
            __check((rounded).toString(), "4")
        }
