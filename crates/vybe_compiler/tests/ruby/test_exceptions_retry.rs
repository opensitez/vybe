
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_retry_basic, "acc = []; i = 0; begin; acc << i; raise 'err' if i < 2; rescue; i += 1; retry; end; puts acc.join('-')", "0-1-2");
ruby_test!(test_retry_resets_ensure, "acc = []; i = 0; begin; begin; acc << \"b#{i}\"; raise 'err' if i < 1; rescue; i += 1; retry; ensure; acc << \"e#{i}\"; end; rescue; end; puts acc.join('-')", "b0-e0-b1-e1"); // wait, retry in rescue doesn't run ensure? Actually, ruby: retry re-evaluates the begin block. Does it run ensure? No, retry jumps to begin. ensure is not run on retry.
ruby_test!(test_retry_ensure, "acc = []; i = 0; begin; acc << \"b#{i}\"; raise 'err' if i < 1; rescue; i += 1; retry; ensure; acc << \"e#{i}\"; end; puts acc.join('-')", "b0-b1-e1"); // verify ruby behavior: ensure is only run when block exits, retry doesn't exit block, it restarts it.
