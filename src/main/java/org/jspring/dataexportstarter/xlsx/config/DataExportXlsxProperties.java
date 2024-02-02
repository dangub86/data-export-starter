package org.jspring.dataexportstarter.xlsx.config;

import org.springframework.boot.context.properties.ConfigurationProperties;


@ConfigurationProperties(prefix = "spring.export.xlsx")
public class DataExportXlsxProperties {

    private String templatePath;

    public DataExportXlsxProperties(String templatePath) {
        this.templatePath = templatePath;
    }

    public String getTemplatePath() {
        return templatePath != null ? templatePath : "src/main/resources/templates/xlsx/model.xls";
    }

    public void setTemplatePath(String templatePath) {
        this.templatePath = templatePath;
    }

}
