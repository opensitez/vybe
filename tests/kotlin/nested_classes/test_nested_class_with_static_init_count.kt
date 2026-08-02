// vybe-test: kotlin/nested_classes/test_nested_class_with_static_init_count
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Counter {
            class Item {
                companion object { var count = 0 }
            }

            init {
                Item.count += 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Counter()
            Counter()
            __check((Counter.Item.count).toString(), "2")
        }
