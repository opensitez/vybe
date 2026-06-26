//! Exception stack traces: catch (e, st), typed on-clauses with st, finally, rethrow.

dart_cases! {
    catch_with_stack_trace_is_non_null => {
        r#"void main() {
  try {
    throw 'boom';
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    catch_stack_trace_is_stacktrace_type => {
        r#"void main() {
  try {
    throw 'x';
  } catch (e, st) {
    print(st is StackTrace);
  }
}"#,
        ["true"]
    };

    catch_stack_trace_to_string_nonempty => {
        r#"void main() {
  try {
    throw 'trace me';
  } catch (e, st) {
    print(st.toString().isNotEmpty);
  }
}"#,
        ["true"]
    };

    catch_stack_trace_available_for_integer_throw => {
        r#"void main() {
  try {
    throw 404;
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    catch_stack_trace_available_for_exception_object => {
        r#"void main() {
  try {
    throw Exception('fail');
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    catch_stack_trace_from_nested_function => {
        r#"void boom() {
  throw 'nested';
}
void main() {
  try {
    boom();
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    catch_stack_trace_from_method_on_class => {
        r#"class Worker {
  void run() {
    throw 'method';
  }
}
void main() {
  try {
    Worker().run();
  } catch (e, st) {
    print(st is StackTrace);
  }
}"#,
        ["true"]
    };

    catch_stack_trace_in_loop_on_throw => {
        r#"void main() {
  try {
    for (var i = 0; i < 3; i++) {
      if (i == 1) throw 'loop';
    }
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    catch_stack_trace_after_conditional_throw => {
        r#"void main() {
  var fail = true;
  try {
    if (fail) throw 'cond';
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    catch_stack_trace_preserves_error_value => {
        r#"void main() {
  try {
    throw 'msg';
  } catch (e, st) {
    print(e);
    print(st != null);
  }
}"#,
        ["msg", "true"]
    };

    on_format_exception_catch_includes_stack_trace => {
        r#"void main() {
  try {
    throw FormatException('bad');
  } on FormatException catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    on_format_exception_stack_trace_is_stacktrace => {
        r#"void main() {
  try {
    throw FormatException('fmt');
  } on FormatException catch (e, st) {
    print(st is StackTrace);
  }
}"#,
        ["true"]
    };

    on_format_exception_catch_reads_message => {
        r#"void main() {
  try {
    throw FormatException('invalid token');
  } on FormatException catch (e, st) {
    print(e.message);
    print(st != null);
  }
}"#,
        ["invalid token", "true"]
    };

    on_format_exception_skipped_for_string_throw => {
        r#"void main() {
  try {
    throw 'plain';
  } on FormatException catch (e, st) {
    print('format');
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    on_format_exception_with_finally_still_present => {
        r#"void main() {
  try {
    throw FormatException('x');
  } on FormatException catch (e, st) {
    print(st != null);
  } finally {
    print('done');
  }
}"#,
        ["true", "done"]
    };

    on_format_exception_from_parse_failure => {
        r#"void main() {
  try {
    int.parse('abc');
  } on FormatException catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    on_format_exception_rethrow_outer_gets_stack => {
        r#"void main() {
  try {
    try {
      throw FormatException('inner');
    } on FormatException catch (e, st) {
      print(st != null);
      rethrow;
    }
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true", "true"]
    };

    on_format_exception_without_catch_variable => {
        r#"void main() {
  try {
    throw FormatException('typed');
  } on FormatException {
    print('typed');
  }
}"#,
        ["typed"]
    };

    on_error_catch_includes_stack_trace => {
        r#"void main() {
  try {
    throw RangeError('bounds');
  } on Error catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    on_error_catch_for_state_error => {
        r#"void main() {
  try {
    throw StateError('bad state');
  } on Error catch (e, st) {
    print(st is StackTrace);
  }
}"#,
        ["true"]
    };

    on_error_catch_for_argument_error => {
        r#"void main() {
  try {
    throw ArgumentError('bad arg');
  } on Error catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    on_error_skipped_for_format_exception => {
        r#"void main() {
  try {
    throw FormatException('fmt');
  } on Error catch (e, st) {
    print('error');
  } on Exception catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    on_error_catch_from_failed_assert_pattern => {
        r#"void main() {
  try {
    throw StateError('assert failed');
  } on Error catch (e, st) {
    print(e is StateError);
    print(st != null);
  }
}"#,
        ["true", "true"]
    };

    on_error_with_finally_runs_both => {
        r#"void main() {
  try {
    throw RangeError('r');
  } on Error catch (e, st) {
    print(st != null);
  } finally {
    print('cleanup');
  }
}"#,
        ["true", "cleanup"]
    };

    on_error_rethrow_preserves_stack_in_outer => {
        r#"void main() {
  try {
    try {
      throw ArgumentError('arg');
    } on Error catch (e, st) {
      print(st != null);
      rethrow;
    }
  } on Error catch (e, st) {
    print(st != null);
  }
}"#,
        ["true", "true"]
    };

    on_error_does_not_catch_string_throw => {
        r#"void main() {
  try {
    throw 'string';
  } on Error catch (e, st) {
    print('error');
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    multiple_on_clauses_format_then_range => {
        r#"void main() {
  try {
    throw RangeError('r');
  } on FormatException catch (e, st) {
    print('format');
  } on RangeError catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    multiple_on_clauses_first_match_wins => {
        r#"void main() {
  try {
    throw FormatException('f');
  } on FormatException catch (e, st) {
    print('format');
    print(st != null);
  } on Exception catch (e, st) {
    print('exception');
  }
}"#,
        ["format", "true"]
    };

    multiple_on_clauses_fallback_generic_catch => {
        r#"void main() {
  try {
    throw 'other';
  } on FormatException catch (e, st) {
    print('format');
  } on RangeError catch (e, st) {
    print('range');
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    multiple_on_clauses_exception_supertype => {
        r#"void main() {
  try {
    throw FormatException('x');
  } on RangeError catch (e, st) {
    print('range');
  } on Exception catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    multiple_on_clauses_error_before_exception => {
        r#"void main() {
  try {
    throw StateError('s');
  } on FormatException catch (e, st) {
    print('format');
  } on Error catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    multiple_on_clauses_three_typed_handlers => {
        r#"void main() {
  try {
    throw ArgumentError('a');
  } on FormatException catch (e, st) {
    print('format');
  } on RangeError catch (e, st) {
    print('range');
  } on ArgumentError catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    multiple_on_clauses_none_match_generic => {
        r#"void main() {
  try {
    throw Exception('generic');
  } on FormatException catch (e, st) {
    print('format');
  } on RangeError catch (e, st) {
    print('range');
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    multiple_on_clauses_with_finally => {
        r#"void main() {
  try {
    throw RangeError('r');
  } on FormatException catch (e, st) {
    print('format');
  } on RangeError catch (e, st) {
    print(st != null);
  } finally {
    print('fin');
  }
}"#,
        ["true", "fin"]
    };

    finally_runs_after_catch_with_stack_trace => {
        r#"void main() {
  try {
    throw 'err';
  } catch (e, st) {
    print('catch');
    print(st != null);
  } finally {
    print('finally');
  }
}"#,
        ["catch", "true", "finally"]
    };

    finally_runs_on_success_without_stack_trace => {
        r#"void main() {
  try {
    print('ok');
  } catch (e, st) {
    print('fail');
  } finally {
    print('fin');
  }
}"#,
        ["ok", "fin"]
    };

    finally_runs_after_on_clause_catch => {
        r#"void main() {
  try {
    throw FormatException('f');
  } on FormatException catch (e, st) {
    print(st != null);
  } finally {
    print('done');
  }
}"#,
        ["true", "done"]
    };

    finally_runs_even_when_rethrow_follows => {
        r#"void main() {
  try {
    try {
      throw 'inner';
    } catch (e, st) {
      print(st != null);
      rethrow;
    }
  } catch (e, st) {
    print('outer');
  } finally {
    print('cleanup');
  }
}"#,
        ["true", "outer", "cleanup"]
    };

    finally_after_multiple_on_and_catch => {
        r#"void main() {
  try {
    throw 'plain';
  } on FormatException catch (e, st) {
    print('format');
  } catch (e, st) {
    print(st != null);
  } finally {
    print('end');
  }
}"#,
        ["true", "end"]
    };

    finally_with_computation_in_try_block => {
        r#"void main() {
  try {
    print(2 + 3);
  } finally {
    print('fin');
  }
}"#,
        ["5", "fin"]
    };

    finally_runs_after_error_on_clause => {
        r#"void main() {
  try {
    throw ArgumentError('a');
  } on Error catch (e, st) {
    print('handled');
  } finally {
    print('fin');
  }
}"#,
        ["handled", "fin"]
    };

    finally_nested_try_finally => {
        r#"void main() {
  try {
    try {
      throw 'deep';
    } catch (e, st) {
      print(st != null);
    } finally {
      print('inner');
    }
  } finally {
    print('outer');
  }
}"#,
        ["true", "inner", "outer"]
    };

    rethrow_caught_exception_outer_has_stack => {
        r#"void main() {
  try {
    try {
      throw 'leaf';
    } catch (e, st) {
      print(st != null);
      rethrow;
    }
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true", "true"]
    };

    rethrow_from_on_format_exception => {
        r#"void main() {
  try {
    try {
      throw FormatException('bad');
    } on FormatException catch (e, st) {
      print(st != null);
      rethrow;
    }
  } on FormatException catch (e, st) {
    print('outer');
    print(st != null);
  }
}"#,
        ["true", "outer", "true"]
    };

    rethrow_preserves_error_message_with_stack => {
        r#"void main() {
  try {
    try {
      throw 'preserve';
    } catch (e, st) {
      rethrow;
    }
  } catch (e, st) {
    print(e);
    print(st != null);
  }
}"#,
        ["preserve", "true"]
    };

    rethrow_from_helper_function => {
        r#"void inner() {
  try {
    throw 'helper';
  } catch (e, st) {
    print(st != null);
    rethrow;
  }
}
void main() {
  try {
    inner();
  } catch (e, st) {
    print(e);
  }
}"#,
        ["true", "helper"]
    };

    rethrow_inside_catch_with_inner_finally => {
        r#"void main() {
  try {
    try {
      throw 'x';
    } catch (e, st) {
      print(st != null);
      rethrow;
    } finally {
      print('inner fin');
    }
  } catch (e, st) {
    print('outer');
    print(st != null);
  }
}"#,
        ["true", "inner fin", "outer", "true"]
    };

    catch_without_stack_variable_still_works => {
        r#"void main() {
  try {
    throw 'solo';
  } catch (e) {
    print(e);
  }
}"#,
        ["solo"]
    };

    stack_trace_on_custom_exception_class => {
        r#"class AppError implements Exception {
  final String msg;
  AppError(this.msg);
}
void main() {
  try {
    throw AppError('custom');
  } catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };

    stack_trace_on_error_subclass_unimplemented => {
        r#"void main() {
  try {
    throw UnimplementedError('todo');
  } on Error catch (e, st) {
    print(st != null);
  }
}"#,
        ["true"]
    };
}
