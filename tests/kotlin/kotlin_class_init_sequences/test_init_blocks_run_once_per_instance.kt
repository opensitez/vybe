// vybe-test: kotlin/kotlin_class_init_sequences/test_init_blocks_run_once_per_instance
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Box {
            init { println(1) }
            init { println(2) }
        }

        fun main() {
            Box()
            Box()
        }

