// vybe-test: kotlin/properties/test_lateinit_property_initialized_lazily_and_checked
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Holder {
            lateinit var label: String

            fun initialize(value: String) {
                label = value
            }
        }

        fun main() {
            val holder = Holder()
            try {
                println(holder.label)
            } catch (error: UninitializedPropertyAccessException) {
                println("not_ready")
            }
            holder.initialize("ok")
            println(holder.label)
        }

