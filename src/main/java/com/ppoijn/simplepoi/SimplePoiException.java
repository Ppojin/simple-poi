package com.ppoijn.simplepoi;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class SimplePoiException extends Throwable {
    public SimplePoiException(Throwable t, Code errorCode, String message) {
        super(t);
        errorCode.printLog(message);
    }
}
