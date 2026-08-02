<?php
// vybe-test: php/php_array_chunk_slice_splice_combine/test_php_array_combine_mismatched_lengths_throws
// origin: languages/php/tests/php/test_php_array_chunk_slice_splice_combine.rs
// vybe-test-mode: compile

try {
    @array_combine(["a"], [1, 2]);
} catch (ValueError $e) {
    echo "Mismatched length caught";
}
