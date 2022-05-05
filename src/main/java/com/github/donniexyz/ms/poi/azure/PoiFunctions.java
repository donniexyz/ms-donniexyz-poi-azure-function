package com.github.donniexyz.ms.poi.azure;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.github.donniexyz.lib.poi.PoiMailMerge;
import com.microsoft.azure.functions.*;
import com.microsoft.azure.functions.annotation.AuthorizationLevel;
import com.microsoft.azure.functions.annotation.FunctionName;
import com.microsoft.azure.functions.annotation.HttpTrigger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

import static com.github.donniexyz.ms.poi.azure.PoiProperties.BASEPATH;

/**
 * Azure Functions with HTTP Trigger.
 */
public class PoiFunctions {

    private static final ObjectMapper objectMapper = new ObjectMapper();

    /**
     * This function listens at endpoint "/api/MailMergePoiFunction". Two ways to invoke it using "curl" command in bash:
     * 1. curl -d "HTTP Body" {your host}/api/MailMergePoiFunction
     * 2. curl "--request POST {your host}/api/MailMergePoiFunction?filename=SomeDocument.docx \
     *     --header 'Content-Type: application/json' \
     *     --data-raw '{
     *         "Nama1": "doz",
     *         "Usia1": 13,
     *         "Alamat": {
     *             "line1": "33X-6-3 The Garden",
     *             "line2": "Apartement Complex",
     *             "city": "Kuala Lumpur"
     *         }
     *     }'"
     */
    @FunctionName("MailMergePoiFunction")
    public HttpResponseMessage runMailMergePoiFunction(
            @HttpTrigger(
                    name = "req",
                    methods = {HttpMethod.GET, HttpMethod.POST},
                    authLevel = AuthorizationLevel.ANONYMOUS)
                    HttpRequestMessage<Optional<String>> request,
            final ExecutionContext context) throws IOException {
        context.getLogger().info("Java HTTP MailMergePoiFunction started with headers " + request.getHeaders());


        // Parse query parameter
        final String filename = request.getQueryParameters().get("filename");
        final String bodyJson = request.getBody().orElseThrow(() -> new RuntimeException("RequestBody is null"));

        Path templatePath = Paths.get(BASEPATH + filename);
        File templateFile = templatePath.toFile();
        if (!templateFile.isFile())
            throw new RuntimeException("Template " + filename + " not found");

        //        byte[] bytes = Files.readAllBytes(templatePath);

        Map<String, Object> bodyMap = objectMapper.readValue(bodyJson, LinkedHashMap.class);

        try (XWPFDocument mailMergedXwpfDocument = PoiMailMerge.perform(templatePath, bodyMap);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            mailMergedXwpfDocument.write(out);
            byte[] bytes = out.toByteArray();

            if (bytes.length <= 0) {
                return request.createResponseBuilder(HttpStatus.BAD_REQUEST).body("{\"errorMessage\": \"Invalid\"}").build();
            } else {
                return request.createResponseBuilder(HttpStatus.OK)
                        .header("Content-Disposition", "inline; filename=\"" + filename + "\"")
                        .body(bytes).build();
            }
        }
    }
}
