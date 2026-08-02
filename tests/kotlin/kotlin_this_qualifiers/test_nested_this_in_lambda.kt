// vybe-test: kotlin/kotlin_this_qualifiers/test_nested_this_in_lambda
// origin: languages/kotlin/tests/kotlin/test_kotlin_this_qualifiers.rs

class Box {
            val marker = "box"
            inner class Holder {
                fun show(prefix: String): String {
                    val read = this@Holder
                    return prefix + read.parentName()
                }

                fun parentName(): String {
                    return this@Box.marker
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
            val h = Box().Holder()
            __check((h.show("m=")).toString(), "m=box")
        }
