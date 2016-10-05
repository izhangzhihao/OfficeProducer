import ch.qos.logback.classic.PatternLayout
import ch.qos.logback.core.ConsoleAppender

scan("60 seconds")
appender("stdOut", ConsoleAppender) {
    layout(PatternLayout) {
        pattern = "%d{HH:mm:ss.SSS} [%thread] %-5level %logger{40} - %msg%n"
    }
}

root(INFO, ["stdOut"])