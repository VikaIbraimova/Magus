package ru.d1g.exceptions;

/**
 * Created by A on 09.05.2017.
 */
public class ParserError extends Error {
    public ParserError(String message) {
        super(message);
    }

    public ParserError(String message, Throwable cause) {
        super(message, cause);
    }
}
