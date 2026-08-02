// vybe-test: kotlin/local_classes/test_local_class_with_companion_in_function
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Factory {
                companion object {
                    fun make(v: Int) = Holder(v)
                }
            }
            class Holder(val v: Int)
            __check((Factory.make(9).v).toString(), "9")
        }
