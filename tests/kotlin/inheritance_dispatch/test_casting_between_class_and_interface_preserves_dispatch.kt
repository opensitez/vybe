// vybe-test: kotlin/inheritance_dispatch/test_casting_between_class_and_interface_preserves_dispatch
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface Labeled {
            fun label(): String = "interface"
        }

        open class Base {
            open fun label(): String = "base"
        }

        class Widget : Base(), Labeled {
            override fun label(): String = "widget"
        }

        fun callThroughInterface(value: Labeled): String {
            return value.label()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val root: Base = Widget()
            val viaInterface = root as Labeled
            __check(((root as Widget).label()).toString(), "widget")
            __check((callThroughInterface(viaInterface)).toString(), "widget")
        }
