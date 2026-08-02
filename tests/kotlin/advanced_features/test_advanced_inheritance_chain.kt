// vybe-test: kotlin/advanced_features/test_advanced_inheritance_chain
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

open class Base {
            open fun id(): String {
                return "base"
            }
        }

        open class Mid : Base() {
            override fun id(): String {
                return "mid"
            }
        }

        class Leaf : Mid() {
            override fun id(): String {
                return super.id() + "+leaf"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val l = Leaf()
            __check((l.id()).toString(), "mid+leaf")
        }
