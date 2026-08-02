// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_for_each
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun main() {
            var total = 0
            sequenceOf(1, 2, 3).forEach { total += it }
            println(total)
        }

