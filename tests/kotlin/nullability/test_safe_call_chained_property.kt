// vybe-test: kotlin/nullability/test_safe_call_chained_property
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

class Profile {
            var name: String? = null
            fun label(): String {
                return name ?: "anon"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p: Profile? = Profile()
            __check((p?.label() ?: "none").toString(), "anon")
        }
