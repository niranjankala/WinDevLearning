using System;

namespace WinDev.Logging {
    public interface ILoggerFactory {
        ILogger CreateLogger(Type type);
    }
}