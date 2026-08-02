// vybe-test: kotlin/data_class_copying/test_data_class_copy_with_nested_defaults
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Inner(val id: Int)
        data class Outer(val inner: Inner, val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Outer(Inner(1), "a")
            val b = a.copy(inner = Inner(2))
            __check((a.inner.id).toString(), "1")
            __check((b.inner.id).toString(), "2")
            __check((b.label).toString(), "a")
        }
