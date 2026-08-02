// vybe-test: kotlin/inheritance_dispatch/test_generic_override_dispatches_by_dynamic_type
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base<T> {
            open fun describe(value: T): String = "base:" + value.toString()
        }

        class Text : Base<String>() {
            override fun describe(value: String): String = "text:" + value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Base<String> = Text()
            __check((item.describe("ok")).toString(), "text:ok")
            __check(((item as Text).describe("x")).toString(), "text:x")
        }
