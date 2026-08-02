// vybe-test: kotlin/collections/test_pair_array_destructuring_loop
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val entries = arrayOf(Pair("left", 1), Pair("right", 2))
            var leftTotal = 0
            var rightTotal = 0
            for ((name, value) in entries) {
                if (name == "left") {
                    leftTotal = value
                } else {
                    rightTotal = value
                }
            }
            println(leftTotal)
            println(rightTotal)
        }

