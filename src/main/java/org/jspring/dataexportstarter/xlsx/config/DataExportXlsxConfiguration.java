package org.jspring.dataexportstarter.xlsx.config;

import org.jspring.dataexportstarter.xlsx.domain.CellWrapper;
import org.jspring.dataexportstarter.xlsx.service.XlsxReadingService;
import org.jspring.dataexportstarter.xlsx.service.XlsxWritingService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.autoconfigure.condition.ConditionalOnClass;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
@EnableConfigurationProperties(DataExportXlsxProperties.class)
@ConditionalOnClass(CellWrapper.class)
public class DataExportXlsxConfiguration {

    @Autowired
    private DataExportXlsxProperties properties;

    @Bean
    @ConditionalOnMissingBean
    public XlsxReadingService readingService() {
        return new XlsxReadingService(properties.getTemplatePath());
    }

    @Bean
    @ConditionalOnMissingBean
    public XlsxWritingService writingService() {
        return new XlsxWritingService();
    }

}
