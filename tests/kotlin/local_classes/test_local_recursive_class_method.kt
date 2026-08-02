// vybe-test: kotlin/local_classes/test_local_recursive_class_method
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Counter(val v: Int) {
                fun next(): Int = if (v <= 0) 0 else Counter(v - 1).next() + 1
            }
            __check((Counter(3).next()).toString(), "3")
        }
