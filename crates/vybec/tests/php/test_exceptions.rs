use super::helpers;
use helpers::compile_ok;

#[test] fn try_catch() { compile_ok("<?php try { throw new Exception('oops'); } catch (Exception $e) { echo $e; }"); }
#[test] fn try_finally() { compile_ok("<?php try { echo 'try'; } finally { echo 'finally'; }"); }
#[test] fn try_catch_finally() { compile_ok("<?php try { echo 'try'; } catch (Exception $e) { echo 'catch'; } finally { echo 'finally'; }"); }
#[test] fn throw_expr() { compile_ok("<?php function fail() { throw new Exception('fail'); }"); }
#[test] fn catch_no_var() { compile_ok("<?php try { throw new Exception('x'); } catch (Exception) { echo 'caught'; }"); }
