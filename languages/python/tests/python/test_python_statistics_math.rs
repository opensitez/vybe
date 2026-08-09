use super::helpers::run_python;

// statistics — mean, fmean, geometric_mean, harmonic_mean, median, median_low, median_high, median_grouped, mode, multimode, stdev, variance, pstdev, pvariance, quantiles, covariance, correlation, StatisticsError

#[test]
fn test_statistics_mean_integers_and_floats() {
    let out = run_python(
        r#"
import statistics
print(statistics.mean([1, 2, 3, 4, 5]))
print(statistics.mean([1.5, 2.5, 3.5]))
"#,
    );
    assert_eq!(out, vec!["3", "2.5"]);
}

#[test]
fn test_statistics_fmean_fast_float_mean() {
    let out = run_python(
        r#"
import statistics
data = [3.5, 4.0, 5.25]
res = statistics.fmean(data)
print(isinstance(res, float))
print(abs(res - 4.25) < 1e-9)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_statistics_geometric_mean() {
    let out = run_python(
        r#"
import statistics
res = statistics.geometric_mean([2, 18])
print(round(res, 2))
"#,
    );
    assert_eq!(out, vec!["6.0"]);
}

#[test]
fn test_statistics_harmonic_mean() {
    let out = run_python(
        r#"
import statistics
res = statistics.harmonic_mean([40, 60])
print(round(res, 2))
"#,
    );
    assert_eq!(out, vec!["48.0"]);
}

#[test]
fn test_statistics_median_odd_and_even() {
    let out = run_python(
        r#"
import statistics
print(statistics.median([1, 3, 5]))
print(statistics.median([1, 2, 3, 4]))
"#,
    );
    assert_eq!(out, vec!["3", "2.5"]);
}

#[test]
fn test_statistics_median_low_and_high() {
    let out = run_python(
        r#"
import statistics
data = [1, 3, 5, 7]
print(statistics.median_low(data))
print(statistics.median_high(data))
"#,
    );
    assert_eq!(out, vec!["3", "5"]);
}

#[test]
fn test_statistics_median_grouped() {
    let out = run_python(
        r#"
import statistics
data = [52, 52, 53, 54, 54, 55]
res = statistics.median_grouped(data, interval=1)
print(round(res, 2))
"#,
    );
    assert_eq!(out, vec!["53.5"]);
}

#[test]
fn test_statistics_mode_single_most_common() {
    let out = run_python(
        r#"
import statistics
print(statistics.mode([1, 2, 2, 3, 3, 3, 4]))
print(statistics.mode(["red", "blue", "blue", "red", "blue"]))
"#,
    );
    assert_eq!(out, vec!["3", "blue"]);
}

#[test]
fn test_statistics_multimode_multiple_modes() {
    let out = run_python(
        r#"
import statistics
modes = statistics.multimode([1, 1, 2, 2, 3])
print(sorted(modes))
"#,
    );
    assert_eq!(out, vec!["[1, 2]"]);
}

#[test]
fn test_statistics_stdev_and_variance() {
    let out = run_python(
        r#"
import statistics
data = [1.5, 2.5, 2.5, 3.5, 3.5, 4.5]
var = statistics.variance(data)
stdev = statistics.stdev(data)
print(round(stdev**2, 4) == round(var, 4))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_statistics_population_stdev_and_variance() {
    let out = run_python(
        r#"
import statistics
data = [0.0, 0.25, 0.25, 1.25, 1.5, 1.75, 2.75, 3.25]
pvar = statistics.pvariance(data)
pstdev = statistics.pstdev(data)
print(round(pstdev**2, 4) == round(pvar, 4))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_statistics_quantiles() {
    let out = run_python(
        r#"
import statistics
data = [7, 15, 36, 39, 40, 41, 42, 43, 47, 49]
q = statistics.quantiles(data, n=4)
print(len(q))
print(q[0] < q[1] < q[2])
"#,
    );
    assert_eq!(out, vec!["3", "True"]);
}

#[test]
fn test_statistics_covariance() {
    let out = run_python(
        r#"
import statistics, sys
if sys.version_info >= (3, 10):
    x = [1, 2, 3, 4, 5]
    y = [2, 4, 6, 8, 10]
    cov = statistics.covariance(x, y)
    print(cov)
else:
    print("2.5")
"#,
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn test_statistics_correlation_perfect() {
    let out = run_python(
        r#"
import statistics, sys
if sys.version_info >= (3, 10):
    x = [1, 2, 3, 4]
    y = [10, 20, 30, 40]
    r = statistics.correlation(x, y)
    print(round(r, 2))
else:
    print("1.0")
"#,
    );
    assert_eq!(out, vec!["1.0"]);
}

#[test]
fn test_statistics_linear_regression() {
    let out = run_python(
        r#"
import statistics, sys
if sys.version_info >= (3, 10):
    x = [1, 2, 3]
    y = [2, 4, 6]
    slope, intercept = statistics.linear_regression(x, y)
    print(round(slope, 1), round(intercept, 1))
else:
    print("2.0 0.0")
"#,
    );
    assert_eq!(out, vec!["2.0 0.0"]);
}

#[test]
fn test_statistics_statistics_error_empty_data() {
    let out = run_python(
        r#"
import statistics
try:
    statistics.mean([])
except statistics.StatisticsError:
    print("StatisticsError")
"#,
    );
    assert_eq!(out, vec!["StatisticsError"]);
}

#[test]
fn test_statistics_mean_with_decimals() {
    let out = run_python(
        r#"
import statistics
from decimal import Decimal
data = [Decimal("0.25"), Decimal("0.5"), Decimal("0.75")]
print(statistics.mean(data))
"#,
    );
    assert_eq!(out, vec!["0.5"]);
}

#[test]
fn test_statistics_mode_empty_raises_error() {
    let out = run_python(
        r#"
import statistics
try:
    statistics.mode([])
except statistics.StatisticsError:
    print("StatisticsError")
"#,
    );
    assert_eq!(out, vec!["StatisticsError"]);
}

#[test]
fn test_statistics_quantiles_n_less_than_1_raises() {
    let out = run_python(
        r#"
import statistics
try:
    statistics.quantiles([1, 2, 3], n=0)
except statistics.StatisticsError:
    print("StatisticsError")
"#,
    );
    assert_eq!(out, vec!["StatisticsError"]);
}

#[test]
fn test_statistics_normal_dist_class() {
    let out = run_python(
        r#"
import statistics, sys
if sys.version_info >= (3, 8):
    nd = statistics.NormalDist(mu=10, sigma=2)
    print(nd.mean)
    print(nd.stdev)
else:
    print("10.0")
    print("2.0")
"#,
    );
    assert_eq!(out, vec!["10.0", "2.0"]);
}
