// vybe-test: kotlin/advanced_features/test_companion_object
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

class Factory {
            companion object {
                fun create(): String {
                    return "Instance Created"
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Factory.create()).toString(), "Instance Created")
        }
