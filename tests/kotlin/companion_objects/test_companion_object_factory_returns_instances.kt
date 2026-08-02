// vybe-test: kotlin/companion_objects/test_companion_object_factory_returns_instances
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Widget private constructor(val label: String) {
            companion object {
                fun create(label: String): Widget = Widget(label)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Widget.create("a")
            val second = Widget.create("b")
            __check((first.label).toString(), "a")
            __check((second.label).toString(), "b")
        }
