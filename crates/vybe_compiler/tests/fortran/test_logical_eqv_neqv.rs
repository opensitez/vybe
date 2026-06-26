//! Fortran logical equivalence (.eqv.) and non-equivalence (.neqv.) operators.

use super::helpers;

fortran_cases! {
    eqv_both_true_yields_true => {
        "program t\nprint *, .true. .eqv. .true.\nend program t\n",
        ["true"]
    };

    eqv_true_with_false_yields_false => {
        "program t\nprint *, .true. .eqv. .false.\nend program t\n",
        ["false"]
    };

    eqv_false_with_true_yields_false => {
        "program t\nprint *, .false. .eqv. .true.\nend program t\n",
        ["false"]
    };

    eqv_both_false_yields_true => {
        "program t\nprint *, .false. .eqv. .false.\nend program t\n",
        ["true"]
    };

    neqv_both_true_yields_false => {
        "program t\nprint *, .true. .neqv. .true.\nend program t\n",
        ["false"]
    };

    neqv_true_with_false_yields_true => {
        "program t\nprint *, .true. .neqv. .false.\nend program t\n",
        ["true"]
    };

    neqv_false_with_true_yields_true => {
        "program t\nprint *, .false. .neqv. .true.\nend program t\n",
        ["true"]
    };

    neqv_both_false_yields_false => {
        "program t\nprint *, .false. .neqv. .false.\nend program t\n",
        ["false"]
    };

    if_eqv_true_true_prints_then_branch => {
        "program t\nif (.true. .eqv. .true.) then\nprint *, \"match\"\nelse\nprint *, \"mismatch\"\nend if\nend program t\n",
        ["match"]
    };

    if_eqv_true_false_prints_else_branch => {
        "program t\nif (.true. .eqv. .false.) then\nprint *, \"match\"\nelse\nprint *, \"mismatch\"\nend if\nend program t\n",
        ["mismatch"]
    };

    if_neqv_differing_values_prints_then_branch => {
        "program t\nif (.true. .neqv. .false.) then\nprint *, \"diff\"\nelse\nprint *, \"same\"\nend if\nend program t\n",
        ["diff"]
    };

    if_neqv_matching_values_prints_else_branch => {
        "program t\nif (.false. .neqv. .false.) then\nprint *, \"diff\"\nelse\nprint *, \"same\"\nend if\nend program t\n",
        ["same"]
    };

    if_eqv_with_stored_variables_prints_yes => {
        "program t\nlogical :: p, q\np = .true.\nq = .true.\nif (p .eqv. q) then\nprint *, \"yes\"\nelse\nprint *, \"no\"\nend if\nend program t\n",
        ["yes"]
    };

    if_neqv_with_stored_variables_prints_diff => {
        "program t\nlogical :: p, q\np = .true.\nq = .false.\nif (p .neqv. q) then\nprint *, \"diff\"\nelse\nprint *, \"same\"\nend if\nend program t\n",
        ["diff"]
    };

    if_eqv_in_elseif_selects_equal_arm => {
        "program t\nif (.true. .neqv. .true.) then\nprint *, \"first\"\nelse if (.true. .eqv. .false.) then\nprint *, \"second\"\nelse\nprint *, \"third\"\nend if\nend program t\n",
        ["third"]
    };

    if_neqv_nested_inside_eqv_picks_else => {
        "program t\nif ((.true. .neqv. .false.) .eqv. .false.) then\nprint *, \"then\"\nelse\nprint *, \"else\"\nend if\nend program t\n",
        ["else"]
    };

    eqv_tt_and_eqv_ff_yields_true => {
        "program t\nprint *, (.true. .eqv. .true.) .and. (.false. .eqv. .false.)\nend program t\n",
        ["true"]
    };

    eqv_tf_or_eqv_tt_yields_true => {
        "program t\nprint *, (.true. .eqv. .false.) .or. (.true. .eqv. .true.)\nend program t\n",
        ["true"]
    };

    neqv_tt_and_neqv_ff_yields_true => {
        "program t\nprint *, (.true. .neqv. .true.) .and. (.false. .neqv. .false.)\nend program t\n",
        ["true"]
    };

    eqv_ff_or_neqv_tt_yields_true => {
        "program t\nprint *, (.false. .eqv. .false.) .or. (.true. .neqv. .true.)\nend program t\n",
        ["true"]
    };

    eqv_true_and_not_neqv_false_yields_true => {
        "program t\nprint *, (.true. .eqv. .true.) .and. .not. (.true. .neqv. .false.)\nend program t\n",
        ["false"]
    };

    neqv_tf_or_eqv_ff_yields_true => {
        "program t\nprint *, (.true. .neqv. .false.) .or. (.false. .eqv. .false.)\nend program t\n",
        ["true"]
    };

    three_way_eqv_chain_all_true_operands => {
        "program t\nprint *, .true. .eqv. .true. .eqv. .true.\nend program t\n",
        ["true"]
    };

    neqv_or_eqv_mixed_with_parentheses => {
        "program t\nprint *, (.true. .neqv. .false.) .or. (.true. .eqv. .false.)\nend program t\n",
        ["true"]
    };

    eqv_tt_result_eqv_eqv_ff_result => {
        "program t\nprint *, (.true. .eqv. .true.) .eqv. (.false. .eqv. .false.)\nend program t\n",
        ["true"]
    };

    eqv_tf_result_eqv_eqv_ft_result => {
        "program t\nprint *, (.true. .eqv. .false.) .eqv. (.false. .eqv. .true.)\nend program t\n",
        ["true"]
    };

    eqv_tt_result_neqv_eqv_tf_result => {
        "program t\nprint *, (.true. .eqv. .true.) .neqv. (.true. .eqv. .false.)\nend program t\n",
        ["true"]
    };

    two_neqv_tf_results_are_eqv => {
        "program t\nprint *, (.true. .neqv. .false.) .eqv. (.false. .neqv. .true.)\nend program t\n",
        ["true"]
    };

    comparing_parallel_eqv_expressions_print_both => {
        "program t\nprint *, .true. .eqv. .true.\nprint *, .false. .eqv. .false.\nend program t\n",
        ["true", "true"]
    };

    eqv_of_two_eqv_results_matches_when_inputs_match => {
        "program t\nprint *, (.true. .eqv. .false.) .eqv. (.false. .eqv. .true.)\nend program t\n",
        ["true"]
    };

    eqv_via_logical_variables_true_true => {
        "program t\nlogical :: a = .true., b = .true.\nprint *, a .eqv. b\nend program t\n",
        ["true"]
    };

    eqv_via_logical_variables_false_false => {
        "program t\nlogical :: a = .false., b = .false.\nprint *, a .eqv. b\nend program t\n",
        ["true"]
    };

    neqv_via_logical_variables_true_false => {
        "program t\nlogical :: a = .true., b = .false.\nprint *, a .neqv. b\nend program t\n",
        ["true"]
    };

    assign_eqv_result_to_variable_and_print => {
        "program t\nlogical :: r\nr = .true. .eqv. .false.\nprint *, r\nend program t\n",
        ["false"]
    };

    assign_neqv_result_to_variable_and_print => {
        "program t\nlogical :: r\nr = .false. .neqv. .true.\nprint *, r\nend program t\n",
        ["true"]
    };

    eqv_with_not_wrapped_operand => {
        "program t\nprint *, .true. .eqv. .not. .false.\nend program t\n",
        ["true"]
    };

    neqv_with_not_wrapped_operand => {
        "program t\nprint *, .false. .neqv. .not. .true.\nend program t\n",
        ["true"]
    };

    if_double_eqv_comparison_nested_prints_match => {
        "program t\nif ((.true. .eqv. .true.) .eqv. (.false. .eqv. .false.)) then\nprint *, \"aligned\"\nelse\nprint *, \"skewed\"\nend if\nend program t\n",
        ["aligned"]
    };

    print_eqv_and_neqv_of_same_pair => {
        "program t\nprint *, .true. .eqv. .false.\nprint *, .true. .neqv. .false.\nend program t\n",
        ["false", "true"]
    };

    eqv_commutes_operand_order => {
        "program t\nprint *, .false. .eqv. .true.\nend program t\n",
        ["false"]
    };

    neqv_commutes_operand_order => {
        "program t\nprint *, .false. .neqv. .true.\nend program t\n",
        ["true"]
    };

    eqv_results_differ_when_one_side_flips => {
        "program t\nprint *, (.true. .eqv. .true.) .eqv. (.true. .eqv. .false.)\nend program t\n",
        ["false"]
    };

    neqv_results_match_when_both_pairs_differ => {
        "program t\nprint *, (.true. .neqv. .false.) .eqv. (.false. .neqv. .true.)\nend program t\n",
        ["true"]
    };

    chained_eqv_and_or_with_ff_neqv_tt => {
        "program t\nprint *, (.false. .eqv. .false.) .and. (.true. .neqv. .true.)\nend program t\n",
        ["false"]
    };

    chained_neqv_and_eqv_with_parentheses => {
        "program t\nprint *, ((.true. .neqv. .false.) .and. (.false. .eqv. .true.)) .eqv. .true.\nend program t\n",
        ["true"]
    };

    if_neqv_guard_skips_body_when_values_agree => {
        "program t\nif (.true. .neqv. .true.) then\nprint *, \"run\"\nend if\nprint *, \"done\"\nend program t\n",
        ["done"]
    };

    if_eqv_guard_runs_body_when_values_agree => {
        "program t\nif (.false. .eqv. .false.) then\nprint *, \"run\"\nend if\nprint *, \"done\"\nend program t\n",
        ["run", "done"]
    };

    compare_two_neqv_expressions_with_eqv => {
        "program t\nprint *, (.true. .neqv. .true.) .eqv. (.false. .neqv. .false.)\nend program t\n",
        ["true"]
    };

    compare_two_eqv_expressions_with_neqv => {
        "program t\nprint *, (.true. .eqv. .false.) .neqv. (.false. .eqv. .true.)\nend program t\n",
        ["false"]
    };
}
