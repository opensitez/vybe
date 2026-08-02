// vybe-test: kotlin/nested_classes/test_nested_class_reference_chain
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Factory {
            class Producer {
                class Widget(val label: String)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val widget = Factory.Producer.Widget("gear")
            __check((widget.label).toString(), "gear")
        }
