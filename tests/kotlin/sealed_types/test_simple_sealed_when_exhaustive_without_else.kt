// vybe-test: kotlin/sealed_types/test_simple_sealed_when_exhaustive_without_else
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Result {
            class Ok(val value: Int) : Result()
            class Fail : Result()
        }

        fun describe(result: Result): String {
            return when (result) {
                is Result.Ok -> "ok:" + result.value.toString()
                is Result.Fail -> "fail"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = describe(Result.Ok(3))
            val other = describe(Result.Fail())
            __check((value).toString(), "ok:3")
            __check((other).toString(), "fail")
        }
