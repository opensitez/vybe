use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    goto_forward_jump => {
        body: r#"
            printf("before\n");
            goto end;
            printf("skipped\n");
            end:
            printf("after\n");
            return 0;
        "#,
        expect: ["before", "after"]
    },
    goto_backward_jump_loop => {
        body: r#"
            int i = 0;
            loop:
            if (i < 3) {
                printf("%d\n", i);
                i++;
                goto loop;
            }
            return 0;
        "#,
        expect: ["0", "1", "2"]
    },
    goto_skip_initialization => {
        body: r#"
            int x = 10;
            if (x > 5) goto done;
            x = 0;
            done:
            printf("%d\n", x);
            return 0;
        "#,
        expect: ["10"]
    },
    goto_multiple_labels => {
        body: r#"
            int n = 2;
            if (n == 1) goto one;
            if (n == 2) goto two;
            goto end;
            one:
            printf("one\n");
            goto end;
            two:
            printf("two\n");
            end:
            return 0;
        "#,
        expect: ["two"]
    },
    goto_nested_block_exit => {
        body: r#"
            int i;
            for (i = 0; i < 5; i++) {
                if (i == 3) goto done;
                printf("%d\n", i);
            }
            done:
            printf("done\n");
            return 0;
        "#,
        expect: ["0", "1", "2", "done"]
    },
    goto_error_handling_pattern => {
        body: r#"
            int ok = 1;
            if (!ok) goto cleanup;
            printf("working\n");
            ok = 0;
            if (!ok) goto cleanup;
            printf("unreachable\n");
            cleanup:
            printf("cleanup\n");
            return 0;
        "#,
        expect: ["working", "cleanup"]
    }
}
