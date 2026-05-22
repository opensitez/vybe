! Statistical Analysis — descriptive stats, sorting, histogram
! Covers: allocatable arrays, optional arguments, elemental functions,
!         character handling, namelist I/O, internal subprograms,
!         select type, associate construct, where/forall.

module stats_module
    implicit none
    private

    integer, parameter :: dp = kind(1.0d0)

    public :: dp, descriptive_stats, histogram, sort_array, percentile

    type, public :: stats_result
        real(dp) :: mean
        real(dp) :: variance
        real(dp) :: std_dev
        real(dp) :: median
        real(dp) :: skewness
        real(dp) :: kurtosis
        real(dp) :: minimum
        real(dp) :: maximum
        real(dp) :: range
        integer  :: n
    contains
        procedure :: print => print_stats
    end type stats_result

contains

    subroutine descriptive_stats(data, result, weights)
        real(dp), intent(in)           :: data(:)
        type(stats_result), intent(out) :: result
        real(dp), intent(in), optional :: weights(:)

        real(dp), allocatable :: w(:), sorted(:)
        real(dp) :: wsum, m2, m3, m4, dev
        integer  :: n, i

        n = size(data)
        result%n = n

        ! Weights
        allocate(w(n))
        if (present(weights)) then
            w = weights / sum(weights)
        else
            w = 1.0_dp / n
        end if

        ! Mean
        result%mean = sum(w * data)

        ! Variance, skewness, kurtosis (Welford-style moments)
        m2 = 0.0_dp; m3 = 0.0_dp; m4 = 0.0_dp
        do i = 1, n
            dev = data(i) - result%mean
            m2 = m2 + w(i) * dev**2
            m3 = m3 + w(i) * dev**3
            m4 = m4 + w(i) * dev**4
        end do

        result%variance = m2
        result%std_dev  = sqrt(m2)
        result%skewness = m3 / (m2**1.5_dp + 1.0e-15_dp)
        result%kurtosis = m4 / (m2**2 + 1.0e-15_dp) - 3.0_dp  ! excess kurtosis

        ! Min/max/range
        result%minimum = minval(data)
        result%maximum = maxval(data)
        result%range   = result%maximum - result%minimum

        ! Median (requires sorted copy)
        allocate(sorted(n))
        sorted = data
        call sort_array(sorted)
        result%median = percentile(sorted, 50.0_dp)

        deallocate(w, sorted)
    end subroutine descriptive_stats

    ! In-place quicksort
    recursive subroutine sort_array(arr)
        real(dp), intent(inout) :: arr(:)
        integer  :: n, pivot_idx
        real(dp) :: pivot, tmp
        integer  :: i, j

        n = size(arr)
        if (n <= 1) return

        ! Median-of-three pivot
        pivot_idx = n / 2
        if (arr(1) > arr(pivot_idx)) then
            tmp = arr(1); arr(1) = arr(pivot_idx); arr(pivot_idx) = tmp
        end if
        if (arr(1) > arr(n)) then
            tmp = arr(1); arr(1) = arr(n); arr(n) = tmp
        end if
        if (arr(pivot_idx) > arr(n)) then
            tmp = arr(pivot_idx); arr(pivot_idx) = arr(n); arr(n) = tmp
        end if
        pivot = arr(pivot_idx)

        i = 1; j = n
        do
            do while (arr(i) < pivot); i = i + 1; end do
            do while (arr(j) > pivot); j = j - 1; end do
            if (i >= j) exit
            tmp = arr(i); arr(i) = arr(j); arr(j) = tmp
            i = i + 1; j = j - 1
        end do

        call sort_array(arr(:j))
        call sort_array(arr(j+1:))
    end subroutine sort_array

    ! Percentile from sorted array (linear interpolation)
    pure function percentile(sorted, p) result(val)
        real(dp), intent(in) :: sorted(:)
        real(dp), intent(in) :: p          ! 0..100
        real(dp) :: val
        real(dp) :: pos, frac
        integer  :: lo, hi, n

        n = size(sorted)
        pos  = (p / 100.0_dp) * (n - 1) + 1.0_dp
        lo   = int(pos)
        hi   = min(lo + 1, n)
        frac = pos - lo

        val = sorted(lo) * (1.0_dp - frac) + sorted(hi) * frac
    end function percentile

    ! ASCII histogram
    subroutine histogram(data, nbins, label)
        real(dp),         intent(in)           :: data(:)
        integer,          intent(in)           :: nbins
        character(len=*), intent(in), optional :: label

        real(dp) :: lo, hi, width
        integer, allocatable :: counts(:)
        integer :: i, bin, max_count
        character(len=40) :: bar

        lo = minval(data)
        hi = maxval(data)
        width = (hi - lo) / nbins

        allocate(counts(nbins))
        counts = 0

        do i = 1, size(data)
            bin = min(int((data(i) - lo) / width) + 1, nbins)
            counts(bin) = counts(bin) + 1
        end do

        max_count = maxval(counts)

        if (present(label)) print *, label
        print *, ""

        do i = 1, nbins
            ! Scale bar to 30 chars
            bar = repeat("*", nint(30.0 * counts(i) / max(max_count, 1)))
            print "(f8.3, a, f8.3, a, i5, 2x, a)", &
                lo + (i-1)*width, " -", lo + i*width, " |", counts(i), trim(bar)
        end do
        print *, ""

        deallocate(counts)
    end subroutine histogram

    subroutine print_stats(self)
        class(stats_result), intent(in) :: self
        print "(a, i0)",      "  N         = ", self%n
        print "(a, f12.4)",   "  Mean      = ", self%mean
        print "(a, f12.4)",   "  Std Dev   = ", self%std_dev
        print "(a, f12.4)",   "  Variance  = ", self%variance
        print "(a, f12.4)",   "  Median    = ", self%median
        print "(a, f12.4)",   "  Skewness  = ", self%skewness
        print "(a, f12.4)",   "  Kurtosis  = ", self%kurtosis
        print "(a, f12.4)",   "  Min       = ", self%minimum
        print "(a, f12.4)",   "  Max       = ", self%maximum
        print "(a, f12.4)",   "  Range     = ", self%range
    end subroutine print_stats

end module stats_module


program statistics
    use stats_module
    implicit none

    integer, parameter :: N = 1000
    real(dp), allocatable :: data1(:), data2(:), combined(:)
    type(stats_result) :: res
    real(dp) :: x, pi
    integer  :: i

    pi = 4.0_dp * atan(1.0_dp)

    allocate(data1(N), data2(N), combined(2*N))

    ! Generate pseudo-random normal data using Box-Muller
    ! (using a simple LCG for reproducibility)
    call generate_normal(data1, N, mean=0.0_dp, std=1.0_dp, seed=42)
    call generate_normal(data2, N, mean=5.0_dp, std=2.0_dp, seed=137)

    ! Combined dataset
    combined(1:N)   = data1
    combined(N+1:)  = data2

    ! --- Dataset 1: Standard Normal ---
    print *, "============================================"
    print *, " Dataset 1: N(0,1) — Standard Normal"
    print *, "============================================"
    call descriptive_stats(data1, res)
    call res%print()
    call histogram(data1, 15, "Histogram of N(0,1):")

    ! Percentiles
    call sort_array(data1)
    print "(a)", "  Percentiles:"
    print "(a, f8.4)", "    P5  = ", percentile(data1, 5.0_dp)
    print "(a, f8.4)", "    P25 = ", percentile(data1, 25.0_dp)
    print "(a, f8.4)", "    P75 = ", percentile(data1, 75.0_dp)
    print "(a, f8.4)", "    P95 = ", percentile(data1, 95.0_dp)

    ! --- Dataset 2: N(5,2) ---
    print *, ""
    print *, "============================================"
    print *, " Dataset 2: N(5,2)"
    print *, "============================================"
    call descriptive_stats(data2, res)
    call res%print()

    ! --- Combined ---
    print *, ""
    print *, "============================================"
    print *, " Combined bimodal distribution"
    print *, "============================================"
    call descriptive_stats(combined, res)
    call res%print()
    call histogram(combined, 20, "Histogram of combined:")

    ! --- Where construct: clip outliers ---
    print *, ""
    print *, "=== Clipping outliers beyond 3 sigma ==="
    associate(mu => res%mean, sigma => res%std_dev)
        where (combined < mu - 3.0_dp * sigma)
            combined = mu - 3.0_dp * sigma
        elsewhere (combined > mu + 3.0_dp * sigma)
            combined = mu + 3.0_dp * sigma
        end where
    end associate
    call descriptive_stats(combined, res)
    print "(a, f8.4)", "  New range after clipping: ", res%range

    deallocate(data1, data2, combined)

contains

    subroutine generate_normal(arr, n, mean, std, seed)
        real(dp), intent(out) :: arr(:)
        integer,  intent(in)  :: n, seed
        real(dp), intent(in)  :: mean, std

        real(dp) :: u1, u2, pi_val
        integer  :: i, s

        pi_val = 4.0_dp * atan(1.0_dp)
        s = seed

        do i = 1, n, 2
            ! LCG random numbers in (0,1)
            s  = mod(s * 1664525 + 1013904223, 2**30)
            u1 = real(s, dp) / 2.0_dp**30
            s  = mod(s * 1664525 + 1013904223, 2**30)
            u2 = real(s, dp) / 2.0_dp**30

            ! Box-Muller transform
            arr(i) = mean + std * sqrt(-2.0_dp * log(u1 + 1.0e-15_dp)) * cos(2.0_dp * pi_val * u2)
            if (i + 1 <= n) then
                arr(i+1) = mean + std * sqrt(-2.0_dp * log(u1 + 1.0e-15_dp)) * sin(2.0_dp * pi_val * u2)
            end if
        end do
    end subroutine generate_normal

end program statistics
