package com.ppoijn.simplepoi;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

@Slf4j
enum Code {
    ETC("ETC", "something went wrong!");

    @Getter
    private final String errorCodeStr;
    private final String errorMessage;

    Code(String errorCodeStr, String errorMessage) {
        this.errorCodeStr = errorCodeStr;
        this.errorMessage = errorMessage;
    }

    public void printLog(){
        log.error("{}: {}", errorCodeStr, errorMessage);
    }

    public void printLog(String message){
        log.error("{}: {};{}", errorCodeStr, errorMessage, message);
    }

    @Override
    public String toString(){
        return errorCodeStr;
    }
}
