// vybe-test: kotlin/sealed_types/test_when_on_sealed_interface_in_function_body_with_local_value
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed interface Stage
        class Start : Stage
        class End : Stage

        fun stage_text(stage: Stage): String {
            val value: Stage = stage
            return when (value) {
                is Start -> "start"
                is End -> "end"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val start: Stage = Start()
            val end: Stage = End()
            __check((stage_text(start)).toString(), "start")
            __check((stage_text(end)).toString(), "end")
        }
