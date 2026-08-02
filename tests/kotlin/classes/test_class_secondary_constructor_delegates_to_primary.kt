// vybe-test: kotlin/classes/test_class_secondary_constructor_delegates_to_primary
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Box {
            val value: Int
            val label: String

            constructor(value: Int) : this(value, "default") {
                __check(("secondary").toString(), "secondary")
            }

            constructor(value: Int, label: String) {
                this.value = value
                this.label = label
            }

            fun describe(): String {
                return label + ":" + value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Box(5)
            __check((item.describe()).toString(), "default:5")
        }
