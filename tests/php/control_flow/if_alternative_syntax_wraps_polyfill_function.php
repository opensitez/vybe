<?php
// vybe-test: php/control_flow/if_alternative_syntax_wraps_polyfill_function
// origin: languages/php/tests/php/test_control_flow.rs
// vybe-test-mode: compile

if ( ! function_exists( 'mb_substr' ) ) :
	function mb_substr( $text, $start, $length = null, $encoding = null ) {
		return _mb_substr( $text, $start, $length, $encoding );
	}
endif;
