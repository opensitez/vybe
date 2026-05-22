! Fast Fourier Transform (Cooley-Tukey radix-2)
! Covers: complex numbers, bit manipulation, do concurrent, elemental functions,
!         assumed-shape arrays, optional arguments, character formatting.

module fft_module
    implicit none
    private

    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: PI = 4.0_dp * atan(1.0_dp)

    public :: dp, PI, fft, ifft, power_spectrum, next_power_of_2, &
              hanning_window, hamming_window, blackman_window

contains

    ! Bit-reversal permutation
    pure subroutine bit_reverse(x)
        complex(dp), intent(inout) :: x(:)
        integer :: n, i, j, k
        complex(dp) :: tmp

        n = size(x)
        j = 0
        do i = 1, n - 1
            k = n / 2
            do while (j >= k)
                j = j - k
                k = k / 2
            end do
            j = j + k
            if (i < j) then
                tmp    = x(i + 1)
                x(i + 1) = x(j + 1)
                x(j + 1) = tmp
            end if
        end do
    end subroutine bit_reverse

    ! In-place Cooley-Tukey FFT
    subroutine fft(x, inverse)
        complex(dp), intent(inout) :: x(:)
        logical, intent(in), optional :: inverse

        integer  :: n, stage, stride, half, i, j
        real(dp) :: angle, sign_val
        complex(dp) :: w, wn, tmp

        n = size(x)
        sign_val = -1.0_dp
        if (present(inverse)) then
            if (inverse) sign_val = 1.0_dp
        end if

        call bit_reverse(x)

        stride = 1
        do while (stride < n)
            half   = stride
            stride = stride * 2
            angle  = sign_val * 2.0_dp * PI / stride
            wn     = cmplx(cos(angle), sin(angle), dp)

            do i = 1, n, stride
                w = cmplx(1.0_dp, 0.0_dp, dp)
                do j = 0, half - 1
                    tmp            = w * x(i + j + half)
                    x(i + j + half) = x(i + j) - tmp
                    x(i + j)        = x(i + j) + tmp
                    w = w * wn
                end do
            end do
        end do

        ! Normalize for inverse
        if (present(inverse)) then
            if (inverse) x = x / real(n, dp)
        end if
    end subroutine fft

    subroutine ifft(x)
        complex(dp), intent(inout) :: x(:)
        call fft(x, inverse=.true.)
    end subroutine ifft

    ! Power spectrum (magnitude squared)
    pure function power_spectrum(x) result(ps)
        complex(dp), intent(in) :: x(:)
        real(dp) :: ps(size(x)/2 + 1)
        integer :: i, n
        n = size(x)
        do i = 1, size(ps)
            ps(i) = real(x(i) * conjg(x(i)), dp) / real(n, dp)**2
        end do
    end function power_spectrum

    ! Next power of 2 >= n
    pure function next_power_of_2(n) result(p)
        integer, intent(in) :: n
        integer :: p
        p = 1
        do while (p < n)
            p = p * 2
        end do
    end function next_power_of_2

    ! Window functions (elemental — can be applied to arrays)
    elemental pure function hanning_window(i, n) result(w)
        integer, intent(in) :: i, n
        real(dp) :: w
        w = 0.5_dp * (1.0_dp - cos(2.0_dp * PI * (i - 1) / (n - 1)))
    end function hanning_window

    elemental pure function hamming_window(i, n) result(w)
        integer, intent(in) :: i, n
        real(dp) :: w
        w = 0.54_dp - 0.46_dp * cos(2.0_dp * PI * (i - 1) / (n - 1))
    end function hamming_window

    elemental pure function blackman_window(i, n) result(w)
        integer, intent(in) :: i, n
        real(dp) :: w
        real(dp) :: alpha
        alpha = 0.16_dp
        w = (1.0_dp - alpha) / 2.0_dp &
          - 0.5_dp * cos(2.0_dp * PI * (i - 1) / (n - 1)) &
          + alpha / 2.0_dp * cos(4.0_dp * PI * (i - 1) / (n - 1))
    end function blackman_window

end module fft_module


program fft_demo
    use fft_module
    implicit none

    integer, parameter :: N = 256
    complex(dp), allocatable :: signal(:), spectrum(:)
    real(dp),    allocatable :: ps(:), window(:), freqs(:)
    real(dp) :: dt, fs, df, t, amp1, amp2, freq1, freq2, noise_amp
    integer  :: i, n_fft, peak_bin
    integer, parameter :: indices(*) = [(i, i = 1, N)]

    ! Signal parameters
    dt        = 1.0_dp / 1000.0_dp   ! 1 kHz sample rate
    fs        = 1.0_dp / dt
    freq1     = 50.0_dp               ! 50 Hz component
    freq2     = 120.0_dp              ! 120 Hz component
    amp1      = 1.0_dp
    amp2      = 0.5_dp
    noise_amp = 0.1_dp

    n_fft = next_power_of_2(N)
    allocate(signal(n_fft), spectrum(n_fft), ps(n_fft/2 + 1), &
             window(N), freqs(n_fft/2 + 1))

    ! Generate test signal: two sinusoids + noise
    ! Use Hanning window to reduce spectral leakage
    window = hanning_window(indices, N)

    do i = 1, N
        t = (i - 1) * dt
        ! Pseudo-random noise (simple LCG)
        signal(i) = cmplx( &
            window(i) * (amp1 * sin(2.0_dp * PI * freq1 * t) + &
                         amp2 * sin(2.0_dp * PI * freq2 * t) + &
                         noise_amp * lcg_noise(i)), &
            0.0_dp, dp)
    end do
    ! Zero-pad
    signal(N+1:) = cmplx(0.0_dp, 0.0_dp, dp)

    ! Copy for spectrum computation
    spectrum = signal

    ! Forward FFT
    call fft(spectrum)

    ! Power spectrum
    ps = power_spectrum(spectrum)

    ! Frequency axis
    df = fs / n_fft
    do i = 1, n_fft/2 + 1
        freqs(i) = (i - 1) * df
    end do

    ! Print results
    print *, "=== FFT Demo ==="
    print "(a, i0, a, f6.1, a)", "  Signal: N=", N, " samples at ", fs, " Hz"
    print "(a, f5.1, a, f5.1, a)", "  Components: ", freq1, " Hz (amp=1.0) + ", freq2, " Hz (amp=0.5)"
    print "(a, f4.2, a)", "  Noise amplitude: ", noise_amp, " (Hanning windowed)"
    print *, ""

    ! Find top 5 peaks in power spectrum
    print *, "=== Top spectral peaks ==="
    print "(a6, a10, a14)", "Rank", "Freq (Hz)", "Power"
    print "(a6, a10, a14)", "----", "---------", "-----"
    call print_top_peaks(ps, freqs, n_fft/2 + 1, 5)

    ! Verify round-trip: IFFT(FFT(x)) == x
    call ifft(spectrum)
    print *, ""
    print "(a, es10.3)", "Round-trip error (max |IFFT(FFT(x)) - x|) = ", &
        maxval(abs(real(spectrum(1:N), dp) - real(signal(1:N), dp)))

    ! Compare window functions
    print *, ""
    print *, "=== Window function comparison (first 8 values) ==="
    print "(a12, 3a12)", "Index", "Hanning", "Hamming", "Blackman"
    do i = 1, 8
        print "(i12, 3f12.4)", i, &
            hanning_window(i, N), &
            hamming_window(i, N), &
            blackman_window(i, N)
    end do

    deallocate(signal, spectrum, ps, window, freqs)

contains

    pure function lcg_noise(seed) result(r)
        integer, intent(in) :: seed
        real(dp) :: r
        integer :: s
        s = mod(seed * 1664525 + 1013904223, 2**30)
        r = real(s, dp) / 2.0_dp**29 - 1.0_dp
    end function lcg_noise

    subroutine print_top_peaks(ps, freqs, n, top_n)
        real(dp), intent(in) :: ps(:), freqs(:)
        integer,  intent(in) :: n, top_n
        real(dp), allocatable :: ps_copy(:)
        integer :: i, j, peak_idx
        real(dp) :: peak_val

        allocate(ps_copy(n))
        ps_copy = ps(1:n)

        do i = 1, top_n
            peak_val = maxval(ps_copy)
            peak_idx = maxloc(ps_copy, 1)
            print "(i6, f10.2, es14.4)", i, freqs(peak_idx), peak_val
            ! Suppress this peak and neighbors
            do j = max(1, peak_idx-3), min(n, peak_idx+3)
                ps_copy(j) = 0.0_dp
            end do
        end do

        deallocate(ps_copy)
    end subroutine print_top_peaks

end program fft_demo
