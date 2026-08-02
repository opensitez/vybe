// vybe-test: kotlin/type_casts/test_type_check_on_sealed_subtype
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

sealed class ResultState {
            class Ok(val value: Int) : ResultState()
            class Err(val message: String) : ResultState()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val state: ResultState = ResultState.Ok(7)
            if (state is ResultState.Ok) {
                __check((state.value).toString(), "7")
            }
            val mapped = state as? ResultState.Err
            __check((mapped == null).toString(), "true")
        }
