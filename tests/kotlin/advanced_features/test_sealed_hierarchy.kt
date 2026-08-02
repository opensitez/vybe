// vybe-test: kotlin/advanced_features/test_sealed_hierarchy
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

sealed class Result {
            class Ok(val value: Int) : Result()
            class Error(val message: String) : Result()
        }

        fun format(result: Result): String {
            return when (result) {
                is Result.Ok -> "ok:" + (result.value)
                is Result.Error -> "error:" + (result.message)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val good = Result.Ok(7)
            val bad = Result.Error("bad")
            __check((format(good)).toString(), "ok:7")
            __check((format(bad)).toString(), "error:bad")
        }
