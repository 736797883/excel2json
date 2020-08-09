package com.zyx.excel2json.service;

import org.springframework.web.multipart.MultipartFile;

public interface Excel2JsonService {

    void excel2JSon(MultipartFile file);
}
