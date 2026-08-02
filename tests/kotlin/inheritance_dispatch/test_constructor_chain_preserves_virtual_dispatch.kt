// vybe-test: kotlin/inheritance_dispatch/test_constructor_chain_preserves_virtual_dispatch
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base(value: Int) {
            init {
                if (value < 0) {
                    println("bad")
                }
            }
        }

        class Child(value: Int) : Base(value) {
            init {
                println("child")
            }
        }

        fun main() {
            Child(3)
        }

