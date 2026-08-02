// vybe-test: kotlin/operators/test_comparison_operator_custom_type
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Version(val major: Int, val minor: Int) {
            operator fun compareTo(other: Version): Int {
                if (major != other.major) {
                    return major - other.major
                }
                return minor - other.minor
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Version(1, 4)
            val b = Version(2, 0)
            val c = Version(1, 2)
            __check((a < b).toString(), "true")
            __check((a > c).toString(), "true")
            __check((a == c).toString(), "false")
        }
