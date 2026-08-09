//! Throw, break, and continue inside switch — distinct control-flow behaviors only.

crate::js_cases! {
    switch_matched_case_throw_caught_by_try => {
        r#"let o=[];try{switch(1){case 1:throw new Error("hit");default:o.push("d");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["hit"]
    };

    switch_default_throw_when_no_case_matches => {
        r#"let o=[];try{switch(99){case 1:throw new Error("case");default:throw new Error("default");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["default"]
    };

    switch_fallthrough_reaches_later_throwing_case => {
        r#"let o=[];try{switch(1){case 1:case 2:throw new Error("fall");default:o.push("d");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["fall"]
    };

    switch_break_in_case_prevents_throw => {
        r#"let o=[];try{switch(1){case 1:o.push("a");break;throw new Error("no");default:o.push("d");}}catch(e){o.push("err");}console.log(o.join(","));"#,
        ["a"]
    };

    switch_throw_string_primitive_reason => {
        r#"let o=[];try{switch("x"){case "x":throw "str";default:o.push("d");}}catch(e){o.push(typeof e+":"+e);}console.log(o.join(","));"#,
        ["string:str"]
    };

    switch_throw_number_primitive_reason => {
        r#"let o=[];try{switch(0){case 0:throw 42;default:o.push("d");}}catch(e){o.push(String(e));}console.log(o.join(","));"#,
        ["42"]
    };

    switch_nested_inner_throw_caught_by_outer_catch => {
        r#"let o=[];try{switch(1){case 1:try{switch(2){case 2:throw new Error("inner");}}catch(e){o.push("in:"+e.message);}break;}}catch(e){o.push("out");}console.log(o.join(","));"#,
        ["in:inner"]
    };

    switch_nested_outer_catch_catches_inner_switch_throw => {
        r#"let o=[];try{switch(1){case 1:try{switch(2){case 2:throw new Error("up");}}catch(e){throw e;}break;}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["up"]
    };

    switch_inside_for_loop_throw_stops_iteration_in_catch => {
        r#"let o=[];for(let i=0;i<3;i++){try{switch(i){case 1:throw new Error("stop");default:o.push(i);}}catch(e){o.push("c");break;}}console.log(o.join(","));"#,
        ["0,c"]
    };

    switch_discriminant_expression_throw_before_cases => {
        r#"let o=[];function boom(){throw new Error("disc");}try{switch(boom()){case 1:o.push("n");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["disc"]
    };

    switch_case_block_with_local_before_throw => {
        r#"let o=[];try{switch(1){case 1:{const x=2;throw new Error("x"+x);}default:o.push("d");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["x2"]
    };

    switch_multiple_labels_share_throw_body => {
        r#"let o=[];try{switch(2){case 1:case 2:throw new Error("shared");default:o.push("d");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["shared"]
    };

    switch_strict_equality_string_one_not_match_number_one => {
        r#"let o=[];try{switch(1){case "1":throw new Error("str");default:throw new Error("def");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["def"]
    };

    switch_case_assignments_run_before_throw => {
        r#"let o=[];try{switch(1){case 1:o.push("pre");throw new Error("post");default:o.push("d");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["pre,post"]
    };

    switch_continue_in_loop_skips_remaining_cases => {
        r#"let o=[];for(let i=0;i<3;i++){try{switch(i){case 0:o.push("a");continue;case 1:throw new Error("b");default:o.push("c");}}catch(e){o.push(e.message);}}console.log(o.join(","));"#,
        ["a,b,c"]
    };

    switch_labeled_break_exits_before_later_throw => {
        r#"let o=[];outer:for(let i=0;i<2;i++){try{switch(i){case 0:o.push("ok");break outer;case 1:throw new Error("late");}}catch(e){o.push("c");}}console.log(o.join(","));"#,
        ["ok"]
    };

    switch_null_discriminant_matches_null_case => {
        r#"let o=[];try{switch(null){case null:throw new Error("null-case");default:o.push("d");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["null-case"]
    };

    switch_boolean_true_matches_true_case => {
        r#"let o=[];try{switch(true){case true:throw new Error("bool");default:o.push("d");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["bool"]
    };

    switch_rethrow_from_catch_preserves_message => {
        r#"let o=[];try{switch(1){case 1:throw new Error("orig");}}catch(e){try{throw e;}catch(x){o.push(x.message);}}console.log(o.join(","));"#,
        ["orig"]
    };

    switch_finally_runs_after_case_throw => {
        r#"let o=[];try{try{switch(1){case 1:throw new Error("t");}}finally{o.push("f");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["f,t"]
    };

    switch_empty_case_falls_through_to_throw => {
        r#"let o=[];try{switch(1){case 1:case 2:throw new Error("end");default:o.push("d");}}catch(e){o.push(e.message);}console.log(o.join(","));"#,
        ["end"]
    };

    switch_throw_typeerror_subclass => {
        r#"let o=[];try{switch(0){case 0:throw new TypeError("bad op");default:o.push("d");}}catch(e){o.push(e.name+":"+e.message);}console.log(o.join(","));"#,
        ["TypeError:bad op"]
    };

    switch_case_expression_throw_during_eval => {
        r#"let o=[]; function bad() { throw new Error("case_eval"); } try { switch(1) { case bad(): o.push("a"); break; } } catch(e) { o.push(e.message); } console.log(o.join(","));"#,
        ["case_eval"]
    };
}
