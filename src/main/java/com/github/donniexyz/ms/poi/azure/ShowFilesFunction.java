package com.github.donniexyz.ms.poi.azure;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.microsoft.azure.functions.*;
import com.microsoft.azure.functions.annotation.AuthorizationLevel;
import com.microsoft.azure.functions.annotation.FunctionName;
import com.microsoft.azure.functions.annotation.HttpTrigger;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import static com.github.donniexyz.ms.poi.azure.PoiProperties.BASEPATH;

/**
 * Azure Functions with HTTP Trigger.
 */
public class ShowFilesFunction {

    private static final ObjectMapper objectMapper = new ObjectMapper();

    /**
     * This function listens at endpoint "/api/HttpExample". Two ways to invoke it using "curl" command in bash:
     * 1. curl -d "HTTP Body" {your host}/api/HttpExample
     * 2. curl "{your host}/api/HttpExample?name=HTTP%20Query"
     */
    @FunctionName("ShowFilesFunction")
    public HttpResponseMessage run(
            @HttpTrigger(
                    name = "req",
                    methods = {HttpMethod.GET, HttpMethod.POST},
                    authLevel = AuthorizationLevel.ANONYMOUS)
                    HttpRequestMessage<Optional<String>> request,
            final ExecutionContext context) {
        context.getLogger().info("Java ShowFilesFunction started with headers " + request.getHeaders());

        // Parse query parameter
        String jsonResponseBody = null;
        try {
            new File(BASEPATH).mkdirs();
            new File(BASEPATH).mkdirs();
            new File(BASEPATH).mkdirs();
            new File(BASEPATH).mkdirs();
            context.getLogger().info("retrieve files from directory: " + BASEPATH);
            List<File> files = Files.list(Paths.get(BASEPATH))
                    .map(Path::toFile)
                    .filter(File::isFile)
                    .collect(Collectors.toList());

            jsonResponseBody = objectMapper.writeValueAsString(files.stream().map(File::getName).collect(Collectors.toList()));
        } catch (IOException e) {
            e.printStackTrace();
        }

        if (jsonResponseBody == null) {
            return request.createResponseBuilder(HttpStatus.BAD_REQUEST).body("{\"errorMessage\": \"Invalid\"}").build();
        } else {
            return request.createResponseBuilder(HttpStatus.OK).body(jsonResponseBody).build();
        }
    }
}
