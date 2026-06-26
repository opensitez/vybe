//! try/catch, on Type, finally, rethrow, string throws, and custom Exception classes.

dart_cases! {
    try_catch_catches_string_throw => {
        r#"void main() {
  try {
    throw 'boom';
  } catch (e) {
    print('caught');
  }
}"#,
        ["caught"]
    };

    catch_prints_string_error_value => {
        r#"void main() {
  try {
    throw 'network down';
  } catch (e) {
    print(e);
  }
}"#,
        ["network down"]
    };

    try_block_runs_before_catch_on_success => {
        r#"void main() {
  try {
    print('inside');
  } catch (e) {
    print('catch');
  }
  print('after');
}"#,
        ["inside", "after"]
    };

    catch_skipped_when_no_exception => {
        r#"void main() {
  try {
    print('ok');
  } catch (e) {
    print('fail');
  }
}"#,
        ["ok"]
    };

    on_format_exception_typed_catch => {
        r#"void main() {
  try {
    throw FormatException('bad format');
  } on FormatException catch (e) {
    print('format');
  }
}"#,
        ["format"]
    };

    on_type_without_catch_variable => {
        r#"void main() {
  try {
    throw FormatException('bad');
  } on FormatException {
    print('typed');
  }
}"#,
        ["typed"]
    };

    on_type_skipped_for_wrong_exception => {
        r#"void main() {
  try {
    throw 'plain';
  } on FormatException {
    print('format');
  } catch (e) {
    print('generic');
  }
}"#,
        ["generic"]
    };

    on_range_error_matches_type => {
        r#"void main() {
  try {
    throw RangeError('out of bounds');
  } on RangeError {
    print('range');
  }
}"#,
        ["range"]
    };

    on_exception_supertype_catches_subclass => {
        r#"void main() {
  try {
    throw FormatException('detail');
  } on Exception {
    print('exception');
  }
}"#,
        ["exception"]
    };

    finally_runs_after_successful_try => {
        r#"void main() {
  try {
    print('try');
  } finally {
    print('finally');
  }
}"#,
        ["try", "finally"]
    };

    finally_runs_after_caught_exception => {
        r#"void main() {
  try {
    throw 'err';
  } catch (e) {
    print('catch');
  } finally {
    print('finally');
  }
}"#,
        ["catch", "finally"]
    };

    finally_runs_without_catch_clause => {
        r#"void main() {
  try {
    print('body');
  } finally {
    print('cleanup');
  }
}"#,
        ["body", "cleanup"]
    };

    finally_after_on_clause => {
        r#"void main() {
  try {
    throw FormatException('x');
  } on FormatException {
    print('on');
  } finally {
    print('done');
  }
}"#,
        ["on", "done"]
    };

    rethrow_propagates_to_outer_catch => {
        r#"void main() {
  try {
    try {
      throw 'inner';
    } catch (e) {
      rethrow;
    }
  } catch (e) {
    print('outer:$e');
  }
}"#,
        ["outer:inner"]
    };

    rethrow_preserves_original_message => {
        r#"void handler() {
  try {
    throw 'preserve me';
  } catch (e) {
    rethrow;
  }
}
void main() {
  try {
    handler();
  } catch (e) {
    print(e);
  }
}"#,
        ["preserve me"]
    };

    throw_string_from_helper_function => {
        r#"void fail() {
  throw 'helper failed';
}
void main() {
  try {
    fail();
  } catch (e) {
    print(e);
  }
}"#,
        ["helper failed"]
    };

    throw_exception_class_message => {
        r#"void main() {
  try {
    throw Exception('oops');
  } catch (e) {
    print('got exception');
  }
}"#,
        ["got exception"]
    };

    custom_exception_class_caught => {
        r#"class AppException implements Exception {
  final String message;
  AppException(this.message);
}
void main() {
  try {
    throw AppException('custom');
  } catch (e) {
    print('caught custom');
  }
}"#,
        ["caught custom"]
    };

    custom_exception_message_printed => {
        r#"class AppException implements Exception {
  final String message;
  AppException(this.message);
}
void main() {
  try {
    throw AppException('bad input');
  } catch (e) {
    var ex = e as AppException;
    print(ex.message);
  }
}"#,
        ["bad input"]
    };

    custom_exception_thrown_from_method => {
        r#"class ParseError implements Exception {
  final String detail;
  ParseError(this.detail);
}
class Parser {
  int parse(String s) {
    if (s.isEmpty) throw ParseError('empty');
    return int.parse(s);
  }
}
void main() {
  try {
    Parser().parse('');
  } catch (e) {
    var err = e as ParseError;
    print(err.detail);
  }
}"#,
        ["empty"]
    };

    nested_try_inner_catch_outer_continues => {
        r#"void main() {
  try {
    try {
      throw 'inner';
    } catch (e) {
      print('inner:$e');
    }
    print('middle');
  } catch (e) {
    print('outer');
  }
}"#,
        ["inner:inner", "middle"]
    };

    nested_try_outer_catches_rethrown => {
        r#"void main() {
  try {
    try {
      throw 'deep';
    } catch (e) {
      rethrow;
    }
  } catch (e) {
    print('surface:$e');
  }
}"#,
        ["surface:deep"]
    };

    catch_after_multiple_on_clauses => {
        r#"void main() {
  try {
    throw RangeError('r');
  } on FormatException {
    print('format');
  } on RangeError {
    print('range');
  } catch (e) {
    print('other');
  }
}"#,
        ["range"]
    };

    on_clause_then_generic_catch_fallback => {
        r#"void main() {
  try {
    throw 'plain';
  } on FormatException {
    print('format');
  } catch (e) {
    print('fallback:$e');
  }
}"#,
        ["fallback:plain"]
    };

    finally_runs_even_when_catch_handles => {
        r#"void main() {
  try {
    throw 'x';
  } catch (e) {
    print('handled');
  } finally {
    print('always');
  }
}"#,
        ["handled", "always"]
    };

    try_catch_with_computation_in_try => {
        r#"void main() {
  try {
    var n = 3 + 4;
    print(n);
  } catch (e) {
    print('err');
  }
}"#,
        ["7"]
    };

    catch_variable_printed_with_prefix => {
        r#"void main() {
  try {
    throw 'start';
  } catch (e) {
    print('wrapped:$e');
  }
}"#,
        ["wrapped:start"]
    };

    throw_integer_coerced_in_catch => {
        r#"void main() {
  try {
    throw 404;
  } catch (e) {
    print(e);
  }
}"#,
        ["404"]
    };

    exception_in_conditional_throw => {
        r#"void main() {
  var flag = true;
  try {
    if (flag) throw 'conditional';
    print('skip');
  } catch (e) {
    print(e);
  }
}"#,
        ["conditional"]
    };

    on_state_error_catches_state_error => {
        r#"void main() {
  try {
    throw StateError('invalid state');
  } on StateError {
    print('state');
  }
}"#,
        ["state"]
    };

    custom_exception_extends_implements => {
        r#"class DataException implements Exception {
  final int code;
  DataException(this.code);
}
void main() {
  try {
    throw DataException(42);
  } catch (e) {
    var d = e as DataException;
    print(d.code);
  }
}"#,
        ["42"]
    };

    try_finally_no_catch_on_success => {
        r#"void main() {
  try {
    print('run');
  } finally {
    print('end');
  }
}"#,
        ["run", "end"]
    };

    rethrow_from_nested_function => {
        r#"void inner() {
  try {
    throw 'leaf';
  } catch (e) {
    rethrow;
  }
}
void main() {
  try {
    inner();
  } catch (e) {
    print(e);
  }
}"#,
        ["leaf"]
    };

    throw_string_with_interpolation_in_message => {
        r#"void main() {
  var id = 7;
  try {
    throw 'item $id missing';
  } catch (e) {
    print(e);
  }
}"#,
        ["item 7 missing"]
    };

    catch_exception_from_loop => {
        r#"void main() {
  try {
    for (var i = 0; i < 3; i++) {
      if (i == 2) throw 'loop break';
    }
  } catch (e) {
    print(e);
  }
}"#,
        ["loop break"]
    };

    on_argument_error_type => {
        r#"void main() {
  try {
    throw ArgumentError('bad arg');
  } on ArgumentError {
    print('arg');
  }
}"#,
        ["arg"]
    };

    finally_after_successful_computation => {
        r#"void main() {
  try {
    print(10 - 3);
  } finally {
    print('fin');
  }
}"#,
        ["7", "fin"]
    };

    custom_exception_factory_style_throw => {
        r#"class ServiceError implements Exception {
  final String service;
  ServiceError(this.service);
}
void main() {
  try {
    throw ServiceError('auth');
  } catch (e) {
    var s = e as ServiceError;
    print(s.service);
  }
}"#,
        ["auth"]
    };

    try_catch_preserves_code_after_block => {
        r#"void main() {
  try {
    throw 'stop';
  } catch (e) {
    print('caught');
  }
  print('continued');
}"#,
        ["caught", "continued"]
    };

    throw_in_else_branch_caught => {
        r#"void main() {
  var ok = false;
  try {
    if (ok) {
      print('fine');
    } else {
      throw 'not ok';
    }
  } catch (e) {
    print(e);
  }
}"#,
        ["not ok"]
    };
}
