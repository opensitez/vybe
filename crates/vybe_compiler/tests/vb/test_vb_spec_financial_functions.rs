use super::helpers::run_vb;

macro_rules! vb_expr_spec {
    ($name:ident, $expr:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let src = format!(
                r#"
Module Program
    Sub Main()
        Console.WriteLine({})
    End Sub
End Module
"#,
                $expr
            );
            let output = run_vb(&src);
            assert_eq!(output, vec![$expected.to_string()]);
        }
    };
}

vb_expr_spec!(
    financial_spec_pmt_zero_rate,
    "Round(Pmt(0, 10, 1000, 0, 0), 2)",
    "-100"
);
vb_expr_spec!(
    financial_spec_fv_zero_rate,
    "Round(FV(0, 10, -100, 0, 0), 2)",
    "1000"
);
vb_expr_spec!(
    financial_spec_pv_zero_rate,
    "Round(PV(0, 10, -100, 0, 0), 2)",
    "1000"
);
vb_expr_spec!(
    financial_spec_nper_zero_rate,
    "Round(NPer(0, -100, 1000, 0, 0), 2)",
    "10"
);
vb_expr_spec!(
    financial_spec_rate_zero_interest_schedule,
    "Round(Rate(10, -100, 1000, 0, 0, 0.1), 4)",
    "0"
);
vb_expr_spec!(
    financial_spec_ipmt_zero_rate,
    "Round(IPmt(0, 1, 10, 1000, 0, 0), 2)",
    "0"
);
vb_expr_spec!(
    financial_spec_ppmt_zero_rate,
    "Round(PPmt(0, 1, 10, 1000, 0, 0), 2)",
    "-100"
);
vb_expr_spec!(
    financial_spec_sln_first_period,
    "Round(SLN(1000, 100, 9), 2)",
    "100"
);
vb_expr_spec!(
    financial_spec_ddb_first_period,
    "Round(DDB(1000, 100, 5, 1), 2)",
    "400"
);
vb_expr_spec!(
    financial_spec_syd_first_period,
    "Round(SYD(1000, 100, 5, 1), 2)",
    "300"
);
