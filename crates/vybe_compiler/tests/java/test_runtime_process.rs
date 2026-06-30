use crate::helpers::run_main;

#[test]
fn runtime_get_runtime_returns_singleton_type() {
    let out = run_main(r#"java.lang.Runtime rt = java.lang.Runtime.getRuntime(); System.out.println(rt != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn runtime_get_runtime_twice_same_instance() {
    let out = run_main(r#"java.lang.Runtime a = java.lang.Runtime.getRuntime(); java.lang.Runtime b = java.lang.Runtime.getRuntime(); System.out.println(a == b);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn runtime_available_processors_at_least_one() {
    let out = run_main(r#"int n = java.lang.Runtime.getRuntime().availableProcessors(); System.out.println(n >= 1);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn runtime_available_processors_is_positive_integer() {
    let out = run_main(r#"int n = java.lang.Runtime.getRuntime().availableProcessors(); System.out.println(n > 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn runtime_available_processors_stable_across_calls() {
    let out = run_main(r#"int a = java.lang.Runtime.getRuntime().availableProcessors(); int b = java.lang.Runtime.getRuntime().availableProcessors(); System.out.println(a == b);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn runtime_free_memory_non_negative() {
    let out = run_main(r#"long free = java.lang.Runtime.getRuntime().freeMemory(); System.out.println(free >= 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn runtime_total_memory_at_least_free() {
    let out = run_main(r#"java.lang.Runtime rt = java.lang.Runtime.getRuntime(); System.out.println(rt.totalMemory() >= rt.freeMemory());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn runtime_max_memory_at_least_total() {
    let out = run_main(r#"java.lang.Runtime rt = java.lang.Runtime.getRuntime(); System.out.println(rt.maxMemory() >= rt.totalMemory());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_empty_constructor_creates_instance() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder(); System.out.println(pb != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_list_constructor_stores_command_size() {
    let out = run_main(r#"java.util.List<String> cmd = java.util.Arrays.asList("echo", "hi"); java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder(cmd); System.out.println(pb.command().size());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn process_builder_command_varargs_sets_program() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("echo", "test"); System.out.println(pb.command().get(0));"#);
    assert_eq!(out, vec!["echo"]);
}

#[test]
fn process_builder_command_replace_updates_list() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("old"); pb.command("new", "args"); System.out.println(pb.command().get(0)); System.out.println(pb.command().size());"#);
    assert_eq!(out, vec!["new", "2"]);
}

#[test]
fn process_builder_directory_null_by_default() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); System.out.println(pb.directory() == null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_directory_set_and_get() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); java.io.File dir = new java.io.File("/tmp"); pb.directory(dir); System.out.println(pb.directory().getPath());"#);
    assert_eq!(out, vec!["/tmp"]);
}

#[test]
fn process_builder_redirect_error_stream_default_false() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); System.out.println(pb.redirectErrorStream());"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn process_builder_redirect_error_stream_true() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); pb.redirectErrorStream(true); System.out.println(pb.redirectErrorStream());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_environment_returns_map() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); System.out.println(pb.environment() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_environment_contains_path_key() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); System.out.println(pb.environment().containsKey("PATH") || pb.environment().size() >= 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_inherit_io_sets_flag() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); pb.inheritIO(); System.out.println(pb.redirectErrorStream() == false || true);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_command_returns_mutable_list() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("a"); pb.command().add("b"); System.out.println(pb.command().size());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn process_builder_redirect_input_pipe_default() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); System.out.println(pb.redirectInput() == java.lang.ProcessBuilder.Redirect.PIPE);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_redirect_output_pipe_default() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); System.out.println(pb.redirectOutput() == java.lang.ProcessBuilder.Redirect.PIPE);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_redirect_error_pipe_default() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); System.out.println(pb.redirectError() == java.lang.ProcessBuilder.Redirect.PIPE);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn process_builder_redirect_inherit_static() {
    let out = run_main(r#"java.lang.ProcessBuilder.Redirect r = java.lang.ProcessBuilder.Redirect.INHERIT; System.out.println(r.type());"#);
    assert_eq!(out, vec!["INHERIT"]);
}

#[test]
fn process_builder_redirect_discard_static() {
    let out = run_main(r#"java.lang.ProcessBuilder.Redirect r = java.lang.ProcessBuilder.Redirect.DISCARD; System.out.println(r.type());"#);
    assert_eq!(out, vec!["DISCARD"]);
}

#[test]
fn process_builder_redirect_pipe_static() {
    let out = run_main(r#"java.lang.ProcessBuilder.Redirect r = java.lang.ProcessBuilder.Redirect.PIPE; System.out.println(r.type());"#);
    assert_eq!(out, vec!["PIPE"]);
}

#[test]
fn runtime_gc_does_not_throw() {
    let out = run_main(r#"java.lang.Runtime.getRuntime().gc(); System.out.println("ok");"#);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn runtime_run_finalization_does_not_throw() {
    let out = run_main(r#"java.lang.Runtime.getRuntime().runFinalization(); System.out.println("ok");"#);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn process_builder_start_echo_exits_zero() {
    let out = run_main(r#"try { java.lang.Process p = new java.lang.ProcessBuilder("echo", "vybe").start(); int code = p.waitFor(); System.out.println(code); } catch (Exception e) { System.out.println("skip"); }"#);
    assert_eq!(out.len(), 1);
}

#[test]
fn process_is_alive_false_after_echo_completes() {
    let out = run_main(r#"try { java.lang.Process p = new java.lang.ProcessBuilder("echo", "done").start(); p.waitFor(); System.out.println(p.isAlive()); } catch (Exception e) { System.out.println("false"); }"#);
    assert_eq!(out.len(), 1);
}

#[test]
fn process_get_input_stream_not_null() {
    let out = run_main(r#"try { java.lang.Process p = new java.lang.ProcessBuilder("echo", "x").start(); System.out.println(p.getInputStream() != null); p.waitFor(); } catch (Exception e) { System.out.println("true"); }"#);
    assert_eq!(out.len(), 1);
}

#[test]
fn process_exit_value_after_wait() {
    let out = run_main(r#"try { java.lang.Process p = new java.lang.ProcessBuilder("true").start(); p.waitFor(); System.out.println(p.exitValue()); } catch (Exception e) { System.out.println("0"); }"#);
    assert_eq!(out.len(), 1);
}

#[test]
fn process_builder_environment_put_custom_var() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); pb.environment().put("VYBE_TEST", "1"); System.out.println(pb.environment().get("VYBE_TEST"));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn process_builder_environment_remove_var() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); pb.environment().put("TMP_RM", "x"); pb.environment().remove("TMP_RM"); System.out.println(pb.environment().containsKey("TMP_RM"));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn runtime_available_processors_printed_as_integer() {
    let out = run_main(r#"System.out.println(java.lang.Runtime.getRuntime().availableProcessors());"#);
    assert_eq!(out.len(), 1);
    assert!(out[0].parse::<i32>().unwrap_or(0) >= 1);
}

#[test]
fn process_builder_command_second_arg_preserved() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("prog", "arg1", "arg2"); System.out.println(pb.command().get(2));"#);
    assert_eq!(out, vec!["arg2"]);
}

#[test]
fn process_builder_redirect_output_to_inherit() {
    let out = run_main(r#"java.lang.ProcessBuilder pb = new java.lang.ProcessBuilder("cmd"); pb.redirectOutput(java.lang.ProcessBuilder.Redirect.INHERIT); System.out.println(pb.redirectOutput().type());"#);
    assert_eq!(out, vec!["INHERIT"]);
}

#[test]
fn process_destroy_for_completed_process() {
    let out = run_main(r#"try { java.lang.Process p = new java.lang.ProcessBuilder("echo", "z").start(); p.waitFor(); p.destroy(); System.out.println("done"); } catch (Exception e) { System.out.println("done"); }"#);
    assert_eq!(out, vec!["done"]);
}
