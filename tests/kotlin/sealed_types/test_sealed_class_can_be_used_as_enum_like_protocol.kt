// vybe-test: kotlin/sealed_types/test_sealed_class_can_be_used_as_enum_like_protocol
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Status {
            class Success(val message: String) : Status()
            class Failure(val code: Int) : Status()
        }

        fun statusCode(status: Status): Int {
            return when (status) {
                is Status.Success -> 0
                is Status.Failure -> status.code
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((statusCode(Status.Success("ok"))).toString(), "0")
            __check((statusCode(Status.Failure(7))).toString(), "7")
        }
