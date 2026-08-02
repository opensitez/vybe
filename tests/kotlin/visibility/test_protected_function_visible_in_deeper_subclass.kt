// vybe-test: kotlin/visibility/test_protected_function_visible_in_deeper_subclass
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

open class Base {
            protected open fun label(): String = "base"
        }

        open class Mid : Base() {
            override fun label(): String = "mid"
        }

        class Child : Mid() {
            fun text(): String = label()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Child().text()).toString(), "mid")
        }
