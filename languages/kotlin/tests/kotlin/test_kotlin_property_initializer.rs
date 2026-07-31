kotlin_run_cases! {
    test_property_initializer_calls_function => (r#"
        var counter = 0

        fun next(): Int {
            counter = counter + 1
            return counter
        }

        class Node {
            val a: Int = next()
            val b: Int = next()
        }

        fun main() {
            val n = Node()
            println(n.a)
            println(n.b)
            println(counter)
        }
    "#, vec!["1", "2", "2"]),
    test_initializer_evaluates_per_instance => (r#"
        var marker = 0

        fun step(): Int {
            marker = marker + 10
            return marker
        }

        class Token {
            val value = step()
        }

        fun main() {
            val first = Token()
            val second = Token()
            println(first.value)
            println(second.value)
        }
    "#, vec!["10", "20"]),
    test_initializer_with_expression => (r#"
        class Meter {
            val label = "x" + 1.toString() + "y"
        }

        fun main() {
            println(Meter().label)
        }
    "#, vec!["x1y"]),
}
