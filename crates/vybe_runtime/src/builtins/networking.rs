use std::net::{TcpStream, TcpListener, UdpSocket, Shutdown};
use std::io::{BufReader, BufRead, Read, Write};
use std::time::Duration;

/// Native networking handle stored in the Interpreter.
/// Each handle is keyed by a unique i64 ID and referenced
/// from VB.NET objects via the `__socket_id` field.
pub enum NetHandle {
    /// A connected TCP stream (for TcpClient / NetworkStream).
    /// We store the stream for writing and a BufReader over a clone for reading.
    Tcp {
        stream: TcpStream,
        reader: BufReader<TcpStream>,
    },
    /// A bound TCP listener (for TcpListener).
    Listener(TcpListener),
    /// A bound UDP socket (for UdpClient).
    Udp(UdpSocket),
}

impl NetHandle {
    /// Create a new TCP handle by connecting to host:port.
    pub fn connect_tcp(host: &str, port: i32, timeout_ms: Option<u64>) -> std::io::Result<Self> {
        let addr = format!("{}:{}", host, port);
        let stream = if let Some(ms) = timeout_ms {
            let addrs: Vec<std::net::SocketAddr> = addr
                .to_socket_addrs_fallback()?;
            let mut last_err = std::io::Error::new(std::io::ErrorKind::AddrNotAvailable, "no addresses");
            let mut connected = None;
            for a in addrs {
                match TcpStream::connect_timeout(&a, Duration::from_millis(ms)) {
                    Ok(s) => { connected = Some(s); break; }
                    Err(e) => last_err = e,
                }
            }
            connected.ok_or(last_err)?
        } else {
            TcpStream::connect(&addr)?
        };
        let reader = BufReader::new(stream.try_clone()?);
        Ok(NetHandle::Tcp { stream, reader })
    }

    /// Create a TCP handle from an already-connected TcpStream (from accept).
    pub fn from_tcp_stream(stream: TcpStream) -> std::io::Result<Self> {
        let reader = BufReader::new(stream.try_clone()?);
        Ok(NetHandle::Tcp { stream, reader })
    }

    /// Bind a TCP listener.
    pub fn bind_listener(addr: &str, port: i32) -> std::io::Result<Self> {
        let bind_addr = format!("{}:{}", addr, port);
        let listener = TcpListener::bind(&bind_addr)?;
        Ok(NetHandle::Listener(listener))
    }

    /// Bind a UDP socket.
    pub fn bind_udp(port: i32) -> std::io::Result<Self> {
        let bind_addr = format!("0.0.0.0:{}", port);
        let socket = UdpSocket::bind(&bind_addr)?;
        Ok(NetHandle::Udp(socket))
    }

    // ── TCP stream operations ──────────────────────────────────────

    /// Read up to `count` bytes from a TCP stream. Returns bytes read.
    pub fn tcp_read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        match self {
            NetHandle::Tcp { reader, .. } => reader.read(buf),
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a TCP stream")),
        }
    }

    /// Read a single byte.
    pub fn tcp_read_byte(&mut self) -> std::io::Result<Option<u8>> {
        let mut buf = [0u8; 1];
        match self.tcp_read(&mut buf) {
            Ok(1) => Ok(Some(buf[0])),
            Ok(0) => Ok(None), // EOF
            Ok(_) => Ok(None),
            Err(e) => Err(e),
        }
    }

    /// Read a line (up to \n).
    pub fn tcp_read_line(&mut self) -> std::io::Result<Option<String>> {
        match self {
            NetHandle::Tcp { reader, .. } => {
                let mut line = String::new();
                let n = reader.read_line(&mut line)?;
                if n == 0 {
                    Ok(None) // EOF
                } else {
                    // Strip trailing \r\n or \n
                    if line.ends_with('\n') { line.pop(); }
                    if line.ends_with('\r') { line.pop(); }
                    Ok(Some(line))
                }
            }
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a TCP stream")),
        }
    }

    /// Read all remaining data as a string.
    pub fn tcp_read_to_end(&mut self) -> std::io::Result<String> {
        match self {
            NetHandle::Tcp { reader, .. } => {
                let mut buf = String::new();
                reader.read_to_string(&mut buf)?;
                Ok(buf)
            }
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a TCP stream")),
        }
    }

    /// Write bytes to a TCP stream.
    pub fn tcp_write(&mut self, data: &[u8]) -> std::io::Result<usize> {
        match self {
            NetHandle::Tcp { stream, .. } => stream.write(data),
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a TCP stream")),
        }
    }

    /// Write a string followed by \r\n.
    pub fn tcp_write_line(&mut self, text: &str) -> std::io::Result<()> {
        match self {
            NetHandle::Tcp { stream, .. } => {
                stream.write_all(text.as_bytes())?;
                stream.write_all(b"\r\n")?;
                Ok(())
            }
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a TCP stream")),
        }
    }

    /// Flush the TCP stream.
    pub fn tcp_flush(&mut self) -> std::io::Result<()> {
        match self {
            NetHandle::Tcp { stream, .. } => stream.flush(),
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a TCP stream")),
        }
    }

    /// Shutdown TCP stream.
    pub fn tcp_shutdown(&mut self) -> std::io::Result<()> {
        match self {
            NetHandle::Tcp { stream, .. } => stream.shutdown(Shutdown::Both),
            _ => Ok(()),
        }
    }

    /// Set read timeout on TCP stream.
    pub fn set_tcp_timeout(&mut self, timeout_ms: Option<u64>) {
        if let NetHandle::Tcp { stream, .. } = self {
            let _ = stream.set_read_timeout(timeout_ms.map(Duration::from_millis));
            let _ = stream.set_write_timeout(timeout_ms.map(Duration::from_millis));
        }
    }

    // ── TCP listener operations ────────────────────────────────────

    /// Accept a connection. Returns the new TcpStream and peer address.
    pub fn listener_accept(&self) -> std::io::Result<(TcpStream, String)> {
        match self {
            NetHandle::Listener(listener) => {
                let (stream, addr) = listener.accept()?;
                Ok((stream, addr.to_string()))
            }
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a TCP listener")),
        }
    }

    /// Check if a connection is pending (non-blocking).
    pub fn listener_pending(&self) -> std::io::Result<bool> {
        match self {
            NetHandle::Listener(listener) => {
                listener.set_nonblocking(true)?;
                let result = match listener.accept() {
                    Ok(_) => true,
                    Err(e) if e.kind() == std::io::ErrorKind::WouldBlock => false,
                    Err(e) => return Err(e),
                };
                listener.set_nonblocking(false)?;
                Ok(result)
            }
            _ => Ok(false),
        }
    }

    // ── UDP operations ─────────────────────────────────────────────

    /// Send data to a specific address.
    pub fn udp_send_to(&self, data: &[u8], host: &str, port: i32) -> std::io::Result<usize> {
        match self {
            NetHandle::Udp(socket) => {
                let addr = format!("{}:{}", host, port);
                socket.send_to(data, &addr)
            }
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a UDP socket")),
        }
    }

    /// Send data to the connected address.
    pub fn udp_send(&self, data: &[u8]) -> std::io::Result<usize> {
        match self {
            NetHandle::Udp(socket) => socket.send(data),
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a UDP socket")),
        }
    }

    /// Receive data. Returns (data, sender_address).
    pub fn udp_recv(&self, buf_size: usize) -> std::io::Result<(Vec<u8>, String)> {
        match self {
            NetHandle::Udp(socket) => {
                let mut buf = vec![0u8; buf_size];
                let (n, addr) = socket.recv_from(&mut buf)?;
                buf.truncate(n);
                Ok((buf, addr.to_string()))
            }
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a UDP socket")),
        }
    }

    /// Connect UDP socket to a default destination.
    pub fn udp_connect(&self, host: &str, port: i32) -> std::io::Result<()> {
        match self {
            NetHandle::Udp(socket) => {
                let addr = format!("{}:{}", host, port);
                socket.connect(&addr)
            }
            _ => Err(std::io::Error::new(std::io::ErrorKind::InvalidInput, "not a UDP socket")),
        }
    }
}

// Helper trait for DNS resolution fallback
trait ToSocketAddrsFallback {
    fn to_socket_addrs_fallback(&self) -> std::io::Result<Vec<std::net::SocketAddr>>;
}

impl ToSocketAddrsFallback for String {
    fn to_socket_addrs_fallback(&self) -> std::io::Result<Vec<std::net::SocketAddr>> {
        use std::net::ToSocketAddrs;
        let addrs: Vec<_> = self.to_socket_addrs()?.collect();
        if addrs.is_empty() {
            Err(std::io::Error::new(std::io::ErrorKind::AddrNotAvailable, format!("cannot resolve: {}", self)))
        } else {
            Ok(addrs)
        }
    }
}
