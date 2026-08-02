// vybe-test: kotlin/inheritance_dispatch/test_dispatch_in_higher_order_function_argument
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Labelled {
            open fun text(): String = "base"
        }

        class Dynamic : Labelled() {
            override fun text(): String = "dynamic"
        }

        fun mapLabel(value: Labelled, render: (Labelled) -> String): String {
            return render(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Labelled = Dynamic()
            __check((mapLabel(item) { it.text() }).toString(), "dynamic")
            __check((mapLabel(item) { target -> "[" + target.text() + "]" }).toString(), "[dynamic]")
        }
