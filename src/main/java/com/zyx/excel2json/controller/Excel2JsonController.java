package com.zyx.excel2json.controller;

import com.zyx.excel2json.service.Excel2JsonService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping
public class Excel2JsonController {

    private Excel2JsonService excel2JsonService;

    @Autowired
    public Excel2JsonController(Excel2JsonService excel2JsonService){
        this.excel2JsonService = excel2JsonService;
    }

    @PostMapping("/excel2Json")
    public void excel2Json(MultipartFile file){

        if(null != file){
            excel2JsonService.excel2JSon(file);
        }
    }
}
