kotlin_run_cases! {
    test_init_order_simple => (r#"
        class Demo {
            init { println("a") }
            init { println("b") }
        }

        fun main() {
            Demo()
        }
    "#, vec!["a", "b"]),
    test_init_block_with_arguments => (r#"
        class Calc(val base: Int) {
            val offset = base + 1

            init {
                println((base * offset).toString())
            }

            init {
                println((offset - base).toString())
            }
        }

        fun main() {
            Calc(3)
        }
    "#, vec!["12", "1"]),
    test_init_inheritance_order => (r#"
        open class Parent {
            init { println("p") }
        }

        class Child : Parent() {
            init { println("c") }
        }

        fun main() {
            Child()
        }
    "#, vec!["p", "c"]),
}
