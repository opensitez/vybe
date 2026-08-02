// vybe-test: kotlin/lateinit_properties/test_lateinit_access_before_initialization_throws
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Box {
            lateinit var text: String
        }

        fun main() {
            val box = Box()
            val result = try {
                println(box.text)
                "ok"
            } catch (e: UninitializedPropertyAccessException) {
                "uninitialized"
            }
            println(result)
        }

