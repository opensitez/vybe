kotlin_run_cases! {
    test_defaulted_constructor_param => (r#"
        class Person(val name: String, val age: Int = 10) {
            fun describe(): String {
                return name + ":" + age.toString()
            }
        }

        fun main() {
            val a = Person("a")
            val b = Person("b", 20)
            println(a.describe())
            println(b.describe())
        }
    "#, vec!["a:10", "b:20"]),
    test_secondary_constructor_flow => (r#"
        class Counter {
            val value: Int

            constructor(base: Int) {
                value = base
            }

            constructor() : this(0)

            fun isZero(): Boolean {
                return value == 0
            }
        }

        fun main() {
            println(Counter().isZero())
            println(Counter(4).isZero())
        }
    "#, vec!["true", "false"]),
    test_constructor_parameter_property => (r#"
        class Box(val payload: String) {
            fun valueLabel(): String {
                return "<" + payload + ">"
            }
        }

        fun main() {
            val b = Box("x")
            println(b.valueLabel())
        }
    "#, vec!["<x>"]),
}
