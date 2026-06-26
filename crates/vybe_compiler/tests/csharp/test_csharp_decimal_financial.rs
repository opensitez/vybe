//! Decimal financial precision: arithmetic, Round, Truncate, Floor, Ceiling, CompareTo.

use crate::csharp_cases;

csharp_cases! {
    decimal_financial_add_currency_line_items => {
        r#"decimal subtotal=19.99m+4.50m+0.01m; Console.WriteLine(subtotal);"#,
        ["24.50"]
    };

    decimal_financial_subtract_change_due => {
        r#"decimal paid=50.00m; decimal total=37.42m; Console.WriteLine(paid-total);"#,
        ["12.58"]
    };

    decimal_financial_multiply_unit_price_by_qty => {
        r#"decimal unit=12.75m; int qty=3; Console.WriteLine(unit*qty);"#,
        ["38.25"]
    };

    decimal_financial_divide_per_unit_cost => {
        r#"decimal bill=100.00m; decimal seats=6m; Console.WriteLine(bill/seats>16.6m&&bill/seats<16.7m);"#,
        ["True"]
    };

    decimal_financial_modulo_penny_remainder => {
        r#"Console.WriteLine(10.01m%0.10m);"#,
        ["0.01"]
    };

    decimal_financial_tax_rate_application => {
        r#"decimal price=100.00m; decimal rate=0.0825m; Console.WriteLine(price*rate);"#,
        ["8.2500"]
    };

    decimal_financial_compound_two_percent => {
        r#"decimal principal=1000.00m; decimal rate=0.02m; Console.WriteLine(principal*(1m+rate));"#,
        ["1020.00"]
    };

    decimal_financial_tip_fifteen_percent => {
        r#"decimal meal=47.80m; Console.WriteLine(meal*0.15m);"#,
        ["7.170"]
    };

    decimal_financial_split_three_ways => {
        r#"decimal total=10.00m; Console.WriteLine(total/3m>3.33m&&total/3m<3.34m);"#,
        ["True"]
    };

    decimal_financial_discount_percentage => {
        r#"decimal list=250.00m; decimal pct=0.20m; Console.WriteLine(list*(1m-pct));"#,
        ["200.00"]
    };

    decimal_financial_round_to_cents_default => {
        r#"Console.WriteLine(decimal.Round(1.235m,2));"#,
        ["1.24"]
    };

    decimal_financial_round_bankers_down => {
        r#"Console.WriteLine(decimal.Round(1.225m,2));"#,
        ["1.22"]
    };

    decimal_financial_round_bankers_up => {
        r#"Console.WriteLine(decimal.Round(2.225m,2));"#,
        ["2.22"]
    };

    decimal_financial_round_zero_decimals => {
        r#"Console.WriteLine(decimal.Round(9.51m,0));"#,
        ["10"]
    };

    decimal_financial_round_negative_value => {
        r#"Console.WriteLine(decimal.Round(-1.235m,2));"#,
        ["-1.24"]
    };

    decimal_financial_round_midpoint_away_from_zero => {
        r#"Console.WriteLine(decimal.Round(1.235m,2,System.MidpointRounding.AwayFromZero));"#,
        ["1.24"]
    };

    decimal_financial_round_midpoint_to_even => {
        r#"Console.WriteLine(decimal.Round(1.235m,2,System.MidpointRounding.ToEven));"#,
        ["1.24"]
    };

    decimal_financial_round_three_decimal_places => {
        r#"Console.WriteLine(decimal.Round(0.1235m,3));"#,
        ["0.124"]
    };

    decimal_financial_round_preserves_trailing_zero_scale => {
        r#"Console.WriteLine(decimal.Round(3.10m,2).ToString("0.00"));"#,
        ["3.10"]
    };

    decimal_financial_truncate_drops_fraction => {
        r#"Console.WriteLine(decimal.Truncate(9.99m));"#,
        ["9"]
    };

    decimal_financial_truncate_negative => {
        r#"Console.WriteLine(decimal.Truncate(-9.99m));"#,
        ["-9"]
    };

    decimal_financial_truncate_already_integral => {
        r#"Console.WriteLine(decimal.Truncate(42.00m));"#,
        ["42"]
    };

    decimal_financial_truncate_small_fraction => {
        r#"Console.WriteLine(decimal.Truncate(0.001m));"#,
        ["0"]
    };

    decimal_financial_floor_positive => {
        r#"Console.WriteLine(System.Math.Floor(3.7m));"#,
        ["3"]
    };

    decimal_financial_floor_negative => {
        r#"Console.WriteLine(System.Math.Floor(-3.2m));"#,
        ["-4"]
    };

    decimal_financial_ceiling_positive => {
        r#"Console.WriteLine(System.Math.Ceiling(3.2m));"#,
        ["4"]
    };

    decimal_financial_ceiling_negative => {
        r#"Console.WriteLine(System.Math.Ceiling(-3.7m));"#,
        ["-3"]
    };

    decimal_financial_compareto_less => {
        r#"Console.WriteLine(1.2m.CompareTo(1.3m));"#,
        ["-1"]
    };

    decimal_financial_compareto_greater => {
        r#"Console.WriteLine(5.0m.CompareTo(4.9m));"#,
        ["1"]
    };

    decimal_financial_compareto_equal_scale => {
        r#"Console.WriteLine(1.0m.CompareTo(1.00m));"#,
        ["0"]
    };

    decimal_financial_compareto_zero => {
        r#"Console.WriteLine(0m.CompareTo(0.0m));"#,
        ["0"]
    };

    decimal_financial_compareto_negative_values => {
        r#"Console.WriteLine((-2m).CompareTo(-1m));"#,
        ["-1"]
    };

    decimal_financial_vat_inclusive_backout => {
        r#"decimal gross=119.00m; decimal vatRate=0.19m; Console.WriteLine(gross/(1m+vatRate));"#,
        ["100"]
    };

    decimal_financial_margin_calculation => {
        r#"decimal revenue=500m; decimal cost=320m; Console.WriteLine((revenue-cost)/revenue);"#,
        ["0.36"]
    };

    decimal_financial_weighted_average => {
        r#"decimal w1=0.6m; decimal w2=0.4m; decimal p1=10m; decimal p2=20m; Console.WriteLine(w1*p1+w2*p2);"#,
        ["14.0"]
    };

    decimal_financial_penny_allocation_first => {
        r#"decimal total=0.10m; int parts=3; decimal share=decimal.Truncate(total/parts*100m)/100m; Console.WriteLine(share);"#,
        ["0.03"]
    };

    decimal_financial_abs_on_loss => {
        r#"decimal pnl=-125.40m; Console.WriteLine(System.Math.Abs(pnl));"#,
        ["125.40"]
    };

    decimal_financial_max_of_two_quotes => {
        r#"Console.WriteLine(System.Math.Max(12.34m,12.35m));"#,
        ["12.35"]
    };

    decimal_financial_min_of_two_quotes => {
        r#"Console.WriteLine(System.Math.Min(12.34m,12.35m));"#,
        ["12.34"]
    };

    decimal_financial_increment_invoice_total => {
        r#"decimal total=99.99m; total+=0.01m; Console.WriteLine(total);"#,
        ["100.00"]
    };

    decimal_financial_decrement_balance => {
        r#"decimal balance=5.00m; balance-=0.01m; Console.WriteLine(balance);"#,
        ["4.99"]
    };

    decimal_financial_unary_negate_credit => {
        r#"decimal credit=250.75m; Console.WriteLine(-credit);"#,
        ["-250.75"]
    };

    decimal_financial_equality_ignores_trailing_zeros => {
        r#"Console.WriteLine(2.50m==2.5m);"#,
        ["True"]
    };

    decimal_financial_less_than_for_budget_cap => {
        r#"decimal spent=999.99m; decimal cap=1000.00m; Console.WriteLine(spent<cap);"#,
        ["True"]
    };

    decimal_financial_greater_or_equal_payment => {
        r#"decimal due=50.00m; decimal paid=50.00m; Console.WriteLine(paid>=due);"#,
        ["True"]
    };

    decimal_financial_parse_from_string => {
        r#"Console.WriteLine(decimal.Parse("1234.56"));"#,
        ["1234.56"]
    };

    decimal_financial_tostring_fixed_two => {
        r#"Console.WriteLine(7.5m.ToString("F2"));"#,
        ["7.50"]
    };

    decimal_financial_round_tie_to_even_half => {
        r#"Console.WriteLine(decimal.Round(2.5m,0));"#,
        ["2"]
    };

    decimal_financial_round_tie_to_even_one_point_five => {
        r#"Console.WriteLine(decimal.Round(1.5m,0));"#,
        ["2"]
    };

    decimal_financial_interest_simple_one_year => {
        r#"decimal p=5000m; decimal r=0.05m; Console.WriteLine(p+p*r);"#,
        ["5250.00"]
    };

    decimal_financial_amortized_payment_estimate => {
        r#"decimal loan=12000m; decimal months=12m; Console.WriteLine(loan/months);"#,
        ["1000"]
    };

    decimal_financial_cumulative_add_loop => {
        r#"decimal sum=0m; for(int i=1;i<=5;i++) sum+=0.1m; Console.WriteLine(sum);"#,
        ["0.5"]
    };
}
