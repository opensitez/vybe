// vybe-test: kotlin/equality_hashcode/test_custom_equals_override_controls_contract
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

class BadEquals(val value: Int) {
            override fun equals(other: Any?): Boolean {
                if (other !is BadEquals) {
                    return false
                }
                return value == other.value
            }

            override fun hashCode(): Int = value
            override fun toString(): String = "BadEquals(" + value.toString() + ")"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = BadEquals(2)
            val second = BadEquals(2)
            __check((first == second).toString(), "true")
            __check((first.toString()).toString(), "BadEquals(2)")
            __check((first.hashCode()).toString(), "2")
        }
