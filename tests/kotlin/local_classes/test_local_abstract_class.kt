// vybe-test: kotlin/local_classes/test_local_abstract_class
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            abstract class A {
                abstract fun v(): Int
            }
            class B : A() {
                override fun v() = 6
            }
            __check((B().v()).toString(), "6")
        }
