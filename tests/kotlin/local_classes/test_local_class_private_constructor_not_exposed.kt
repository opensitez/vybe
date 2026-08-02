// vybe-test: kotlin/local_classes/test_local_class_private_constructor_not_exposed
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class C private constructor(val v: Int) {
                companion object {
                    fun make(v: Int) = C(v)
                }
            }
            __check((C.make(2).v).toString(), "2")
        }
