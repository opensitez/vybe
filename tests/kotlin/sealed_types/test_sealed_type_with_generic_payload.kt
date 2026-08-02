// vybe-test: kotlin/sealed_types/test_sealed_type_with_generic_payload
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Result<T> {
            class Ok<T>(val value: T) : Result<T>()
            class Error<T>(val reason: String) : Result<T>()
        }

        fun render(value: Result<Int>): String {
            return when (value) {
                is Result.Ok -> "ok:" + value.value.toString()
                is Result.Error -> "err:" + value.reason
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((render(Result.Ok(4))).toString(), "ok:4")
            __check((render(Result.Error("x"))).toString(), "err:x")
        }
