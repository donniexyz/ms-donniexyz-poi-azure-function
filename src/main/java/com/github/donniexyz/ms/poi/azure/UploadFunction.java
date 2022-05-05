package com.github.donniexyz.ms.poi.azure;

import com.microsoft.azure.functions.*;
import com.microsoft.azure.functions.annotation.AuthorizationLevel;
import com.microsoft.azure.functions.annotation.FunctionName;
import com.microsoft.azure.functions.annotation.HttpTrigger;
import org.apache.commons.fileupload.MultipartStream;

import java.io.*;
import java.util.Optional;

import static com.github.donniexyz.ms.poi.azure.PoiProperties.BASEPATH;

/**
 * Azure Functions with HTTP Trigger.
 */
public class UploadFunction {


    /**
     * This function listens at endpoint "/api/HttpExample". Two ways to invoke it using "curl" command in bash:
     * 1. curl -d "HTTP Body" {your host}/api/HttpExample
     * 2. curl "{your host}/api/HttpExample?name=HTTP%20Query"
     */
    @FunctionName("UploadFunction")
    public HttpResponseMessage run(
            @HttpTrigger(
                    name = "req",
                    methods = {HttpMethod.POST},
                    authLevel = AuthorizationLevel.ANONYMOUS,
                    dataType = "binary")
                    HttpRequestMessage<Optional<byte[]>> request,
            final ExecutionContext context) throws IOException {

        context.getLogger().info("Java HTTP UploadFunction started with headers " + request.getHeaders());

        new File(BASEPATH).mkdirs();
        new File(BASEPATH).mkdirs();
        new File(BASEPATH).mkdirs();
        new File(BASEPATH).mkdirs();

        // parse headers
        String contentType = request.getHeaders().get("content-type"); // Get content-type header
        // here the "content-type" must be lower-case
        if (!"multipart/form-data".equals(contentType.split(";")[0]))
            return request.createResponseBuilder(HttpStatus.BAD_REQUEST).build();
        byte[] bs = request.getBody().get();
        InputStream in = new ByteArrayInputStream(bs); // Convert body to an input stream
        String boundary = contentType.split(";")[1].split("=")[1]; // Get boundary from content-type header
        int bufSize = 1024;

        MultipartStream multipartStream = new MultipartStream(in, boundary.getBytes(), bufSize, null); // Using
        // MultipartStream to
        // parse body input
        // stream
        // read files into the specified folder
        boolean nextPart = multipartStream.skipPreamble();

        String fileName = "";

        while (nextPart) {
            String multipartHeader = multipartStream.readHeaders();

            if (multipartHeader.contains("filename")) {
                // parse file and save to disk
                fileName = extractFileName(multipartHeader);
                context.getLogger().info("Processed part with header " + multipartHeader);
                context.getLogger().info("Processed part with filename " + fileName);

                // try with resources
                try (
                        FileOutputStream fos = new FileOutputStream(BASEPATH + fileName);
                ) {
                    context.getLogger().info("Saving file " + BASEPATH + fileName);
                    multipartStream.readBodyData(fos);
                }
            } else {
                // parse text field and do something with it
                context.getLogger().info("Processed text field");
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                multipartStream.readBodyData(baos);
                System.err.println("Field is " + baos.toString());
            }

            nextPart = multipartStream.readBoundary();
        }

        // return response
        context.getLogger().info("Java HTTP file upload ended. Length: " + bs.length);
        return request.createResponseBuilder(HttpStatus.OK).body("Uploaded, fileName:" + fileName + " fileSize: " + bs.length).build();
    }

    // extracts file name from a multipart boundary
    public static String extractFileName(String header) {
        final String FILENAME_PARAMETER = "filename=";
        final int FILENAME_INDEX = header.indexOf(FILENAME_PARAMETER);
        String name = header.substring(FILENAME_INDEX + FILENAME_PARAMETER.length(), header.lastIndexOf("\""));
        String fileName = name.replaceAll("\"", "").replaceAll(" ", "");

        return fileName;
    }
}
