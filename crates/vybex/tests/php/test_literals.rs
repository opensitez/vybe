use super::helpers::compile_ok;

#[test] fn int_literal() { compile_ok("<?php $x = 42;"); }
#[test] fn float_literal() { compile_ok("<?php $x = 3.14;"); }
#[test] fn string_single() { compile_ok("<?php $x = 'hello';"); }
#[test] fn string_double() { compile_ok("<?php $x = \"hello\";"); }
#[test] fn bool_true() { compile_ok("<?php $x = true;"); }
#[test] fn bool_false() { compile_ok("<?php $x = false;"); }
#[test] fn null_literal() { compile_ok("<?php $x = null;"); }
#[test] fn array_indexed() { compile_ok("<?php $x = [1, 2, 3];"); }
#[test] fn array_assoc() { compile_ok("<?php $x = ['name' => 'John', 'age' => 30];"); }
#[test] fn array_empty() { compile_ok("<?php $x = [];"); }
#[test] fn array_nested() { compile_ok("<?php $x = [[1,2],[3,4]];"); }
#[test] fn array_mixed_keys() { compile_ok("<?php $x = [0 => 'a', 'key' => 'b', 1 => 'c'];"); }
#[test] fn const_decl() { compile_ok("<?php const FOO = 42;"); }
