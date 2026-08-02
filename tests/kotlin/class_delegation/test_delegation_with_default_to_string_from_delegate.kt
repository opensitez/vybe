// vybe-test: kotlin/class_delegation/test_delegation_with_default_to_string_from_delegate
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Named { fun title(): String }

        class NamedImpl : Named {
            override fun title() = "named"
            override fun toString() = "impl"
        }

        class NamedProxy(delegate: Named) : Named by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = NamedProxy(NamedImpl())
            __check((value.title()).toString(), "named")
            __check((value.toString()).toString(), "impl")
        }
