import java.io.*;
import java.nio.charset.Charset;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class DocxHelper {

    public static void saveXmlToDocx(File sourceDocxFile, String targetDocxFilePath, String newXmlString) throws IOException {

        ZipFile original = new ZipFile(sourceDocxFile);
        ZipOutputStream outputStream = new ZipOutputStream(new FileOutputStream(targetDocxFilePath));

        Enumeration entries = original.entries();
        while (entries.hasMoreElements()) {

            ZipEntry entry = (ZipEntry)entries.nextElement();
            ZipEntry newEntry = new ZipEntry(entry.getName());

            outputStream.putNextEntry(newEntry);
            if  ("word/document.xml".equalsIgnoreCase(entry.getName())) {

                byte[] bufferOut = new byte[512];
                InputStream in = new ByteArrayInputStream(newXmlString.getBytes(Charset.forName("UTF-8")));
                while (0 < in.available()){
                    int read = in.read(bufferOut);
                    outputStream.write(bufferOut,0,read);
                }
                in.close();
            }
            else{

                byte[] buffer = new byte[(int) entry.getSize()];
                InputStream in = original.getInputStream(entry);
                while (0 < in.available()){
                    int read = in.read(buffer);
                    outputStream.write(buffer,0,read);
                }
                in.close();
            }
            outputStream.closeEntry();
        }

        outputStream.close();
        original.close();
    }

    public static void saveXmlToDocx(String sourceDocxFilePath, String targetDocxFilePath, String newXmlString) throws IOException {

        saveXmlToDocx(new File(sourceDocxFilePath), targetDocxFilePath, newXmlString);
    }

    public static String loadXmlFromDocx(File sourceDocxFile) throws IOException {

        ZipFile zip = new ZipFile(sourceDocxFile);
        InputStream isXml = zip.getInputStream(zip.getEntry("word/document.xml"));
        BufferedReader br = new BufferedReader(new InputStreamReader(isXml, "UTF-8"));

        String part;
        StringBuilder out = new StringBuilder();
        while((part = br.readLine()) != null)
            out.append(part);

        br.close();
        isXml.close();
        zip.close();

        return out.toString();
    }

    public static String loadXmlFromDocx(String sourceDocxFilePath) throws IOException {

        return loadXmlFromDocx(new File(sourceDocxFilePath));
    }

}

package com.company.customXmlParser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class XmlDocument {

    String xmlString; // <- visible inside the nodes (NodeDefault)

    private int xmlStringLength;

    private String xmlDeclaration;

    //-----------------

    private NodeDefault root;

    public XmlDocument(String xmlString) throws Exception {

        this.xmlString = xmlString;
        this.xmlStringLength = xmlString.length();
        root = parse();

        // --- resolve VirtualBlocks (move nodes from body into children of new virtual node)
        ArrayList<NodeDefault> childrenOfVirtualBlock = null;
        NodeDefault newVirtualNode = null;
        int indexOfNewVirtualBlock = -1;

        ArrayList<NodeDefault> bodyChildren = root.children.get(0).children;
        for(int i = 0; i < bodyChildren.size(); i++) {
            NodeDefault childOfBody = bodyChildren.get(i);
            if (childOfBody instanceof NodeSdt && ((NodeSdt) childOfBody).getBindingPropertyName().startsWith("CiProLoopBegin_")) {
                childrenOfVirtualBlock = new ArrayList<>();
                indexOfNewVirtualBlock = i;
            }

            if (childrenOfVirtualBlock != null) {
                childrenOfVirtualBlock.add(childOfBody);
                if (childOfBody instanceof NodeSdt && ((NodeSdt) childOfBody).getBindingPropertyName().equals("CiProLoopEnd")) {
                    newVirtualNode = NodeDefault.createNewVirtualBlock(this, childrenOfVirtualBlock);
                    break;
                }
            }
        }

        if (newVirtualNode != null){
            Iterator<NodeDefault> iterator = newVirtualNode.children.iterator();
            while (iterator.hasNext()) {
                root.children.get(0).children.remove(iterator.next());
            }
            bodyChildren.add(indexOfNewVirtualBlock, newVirtualNode);
        }
        //--------------
    }

    // return index of opening tag <something
    public int indexOfOpeningTag(int indexFrom) {

        int L = indexFrom - 1;
        do {
            if ((L = xmlString.indexOf("<", L + 1)) < 0)
                break;
        } while(xmlString.charAt(L + 1) == '/');

        return L;
    }

    private NodeDefault parse() throws Exception {

        NodeDefault treeToReturn = null;
        Stack stack = new Stack();

        int currentNodeIndex = 0;
        int nextNodeIndex = xmlString.indexOf("<w", 0);
        if (nextNodeIndex >= 0) {

            // resolve xmlDeclaration (everything before root node)
            xmlDeclaration = xmlString.substring(0, nextNodeIndex);

            // parse xmlString and store result in treeToReturn
            while (nextNodeIndex < xmlStringLength && treeToReturn == null) {
                // resolve currentNodeIndex and nextNodeIndex
                currentNodeIndex = nextNodeIndex;

                if ((nextNodeIndex = indexOfOpeningTag(nextNodeIndex + 1)) < 0)
                    nextNodeIndex = xmlStringLength;

                // create instance of new node
                NodeDefault newNode = NodeDefault.createNewNode(this, currentNodeIndex, nextNodeIndex); //

                // two possibilities
                int _nodeTerminatorIndex = indexOfBetween("/>", currentNodeIndex, nextNodeIndex);
                if (_nodeTerminatorIndex > 0) {
                    newNode.setOuterNodeEnd(_nodeTerminatorIndex + "/>".length());
                    stack.peek().children.add(newNode);
                } else {
                    newNode.setInnerNodeBegin(indexOfBetween(">", currentNodeIndex, nextNodeIndex) + 1);
                    stack.push(newNode);
                }

                //
                NodeDefault currentNode = stack.peek();
                int _closingTagIndex = currentNodeIndex;
                while ((_closingTagIndex = indexOfBetween("</", _closingTagIndex + 1, nextNodeIndex)) > 0) {
                    currentNode.setInnerNodeEnd(_closingTagIndex);
                    currentNode.setOuterNodeEnd((_closingTagIndex = indexOfBetween(">", _closingTagIndex, nextNodeIndex)) + 1);

                    if (stack.size() > 1) {
                        stack.pop();
                        stack.peek().children.add(currentNode);
                        currentNode = stack.peek();
                    } else {
                        treeToReturn = stack.pop();
                        break; // <- parsing finished
                    }
                }

            }
        }

        return treeToReturn;
    }

    public String BindTheDataAndGetXml(HashMap dataToInsert, IfDataNodeIsNull bindingAction) {
        StringBuilder resultXmlString = new StringBuilder(xmlDeclaration);

        processDataBinding(root, dataToInsert, resultXmlString, bindingAction);

        return resultXmlString.toString();
    }

    private void processDataBinding(NodeDefault currentSourceElement, HashMap<String, Object> dataToInsert, StringBuilder resultXmlString, IfDataNodeIsNull bindingAction) {

        if (dataToInsert == null)
            return;

        if (currentSourceElement instanceof NodeSdt) {
            String propertyName = ((NodeSdt) currentSourceElement).getBindingPropertyName();
            if (dataToInsert.containsKey(propertyName))
                resultXmlString.append(((NodeSdt) currentSourceElement).BindTheDataAndGetXml((String) dataToInsert.get(propertyName)));
            else if (bindingAction.equals(IfDataNodeIsNull.DontChange))
                resultXmlString.append(currentSourceElement.getOuterXml());
        }

        else if (currentSourceElement instanceof NodeTbl) {
            String propertyName = ((NodeTbl) currentSourceElement).getBindingPropertyName();
            if (dataToInsert.containsKey(propertyName))
                resultXmlString.append(((NodeTbl) currentSourceElement).BindTheDataAndGetXml(dataToInsert.get(propertyName)));
            else if (bindingAction.equals(IfDataNodeIsNull.RemoveBindingOnly))
                resultXmlString.append(((NodeTbl) currentSourceElement).getTheNodeTemplate());
            else if (bindingAction.equals(IfDataNodeIsNull.DontChange))
                resultXmlString.append(currentSourceElement.getOuterXml());
        }

        else if (currentSourceElement instanceof NodeVirtualBlock) {
            String propertyName = ((NodeVirtualBlock) currentSourceElement).getBindingPropertyName();
            if (dataToInsert.containsKey(propertyName))
                resultXmlString.append(((NodeVirtualBlock) currentSourceElement).BindTheDataAndGetXml(dataToInsert.get(propertyName)));
            else if (bindingAction.equals(IfDataNodeIsNull.RemoveBindingOnly))
                resultXmlString.append(((NodeVirtualBlock) currentSourceElement).getTheNodeTemplate());
            else if (bindingAction.equals(IfDataNodeIsNull.DontChange))
                resultXmlString.append(currentSourceElement.getOuterXml());
        }

        else {

            Iterator<NodeDefault> iterator = currentSourceElement.children.iterator();
            if (iterator.hasNext()) {

                resultXmlString.append(currentSourceElement.getOpeningTag());
                while (iterator.hasNext()) {
                    NodeDefault childSourceElement = iterator.next();
                    processDataBinding(childSourceElement, dataToInsert, resultXmlString, bindingAction);
                }
                resultXmlString.append(currentSourceElement.getClosingTag());

            } else {
                resultXmlString.append(currentSourceElement.getOuterXml());
            }
        }
    }

    // helper function visible inside the nodes (NodeDefault)
    int indexOfBetween(String pattern, int startIndex, int endIndex) {

        int _indexOf = xmlString.indexOf(pattern, startIndex);

        if (_indexOf >= endIndex)
            _indexOf = -1;

        return _indexOf;
    }
}

package com.company.customXmlParser;

import java.util.ArrayList;

public class NodeDefault {

    private XmlDocument _SD; // <- sourceDocument

    private String nodeName;

    private int outerNodeBegin;
    private Integer innerNodeBegin;
    private Integer innerNodeEnd;
    private int outerNodeEnd = -1;

    public ArrayList<NodeDefault> children;

    // constructor not public (it is accessible only inside the package)
    NodeDefault(){
    }

    // factory is static public
    static NodeDefault createNewNode(XmlDocument sourceDocument, int currentNodeIndex, int nextNodeIndex) throws Exception {

        String nodeName = resolveNodeName(sourceDocument, currentNodeIndex, nextNodeIndex);

        NodeDefault newNode;
        switch(nodeName){
            case "w:sdt":
                newNode = new NodeSdt();
                break;
            case "w:tbl":
                {
                    int _tbl_closing_index = sourceDocument.xmlString.indexOf("</w:tbl>", currentNodeIndex);
                    int _sdt_index = sourceDocument.indexOfBetween("<w:sdt>", currentNodeIndex, _tbl_closing_index);
                    int _ciPro_TableList_index = sourceDocument.indexOfBetween("<w:tag w:val=\"CiProTable_", _sdt_index, _tbl_closing_index);
                    if (_tbl_closing_index > 0 &&_sdt_index >= 0 && _ciPro_TableList_index >= 0 && _ciPro_TableList_index > _sdt_index)
                        newNode = new NodeTbl();
                    else
                        newNode = new NodeDefault();
                }
                break;
            default:
                newNode = new NodeDefault();
                break;
        }

        newNode.initializeFields(sourceDocument, currentNodeIndex, nodeName);

        return newNode;
    }

    // factory for VirtualBlock
    public static NodeDefault createNewVirtualBlock(XmlDocument sourceDocument, ArrayList<NodeDefault> childrenOfVirtualBlock)
    {
        int firstNodeBegin = childrenOfVirtualBlock.get(0).outerNodeBegin;
        int lastNodeEnd = childrenOfVirtualBlock.get(childrenOfVirtualBlock.size() - 1).outerNodeEnd;

        NodeDefault newNode = new NodeVirtualBlock();
        newNode.initializeFields(sourceDocument, firstNodeBegin, "VirtualBlock");

        newNode.children = childrenOfVirtualBlock;

        newNode.setInnerNodeBegin(firstNodeBegin); // <- for virtual node innerNodeBegin is equal to outerNodeBegin
        newNode.setInnerNodeEnd(lastNodeEnd);
        newNode.setOuterNodeEnd(lastNodeEnd); // <- for virtual node innerNodeEnd is equal to outerNodeEnd

        return newNode;
    }

    // initializer public
    public void initializeFields(XmlDocument sourceDocument, int nodeStartIndex, String nodeName){

        this._SD = sourceDocument;
        this.nodeName = nodeName;

        this.outerNodeBegin = nodeStartIndex;
        this.innerNodeBegin = null;
        this.innerNodeEnd = null;

        this.children = new ArrayList<>();
    }

    // values cannot be changed after outerNodeEnd gets the positive value
    public void setInnerNodeBegin(int idx){ if(outerNodeEnd < 0) innerNodeBegin = idx; }
    public void setInnerNodeEnd(int idx){
        if(outerNodeEnd < 0) innerNodeEnd = idx;
    }
    public void setOuterNodeEnd(int idx){ if(outerNodeEnd < 0) outerNodeEnd = idx; }

    public String getName(){ return nodeName; }

    public String getOpeningTag(){ return _SD.xmlString.substring(outerNodeBegin, innerNodeBegin); }
    public String getClosingTag(){ return _SD.xmlString.substring(innerNodeEnd, outerNodeEnd); }
    public String getInnerXml(){ return _SD.xmlString.substring(innerNodeBegin, innerNodeEnd); }
    public String getOuterXml(){ return _SD.xmlString.substring(outerNodeBegin, outerNodeEnd); }

    //public int getOuterNodeBegin(){ return outerNodeBegin; }
    //public int getInnerNodeBegin(){ return innerNodeBegin; }
    //public int getInnerNodeEnd(){ return innerNodeEnd; }
    //public int getouterNodeEnd(){ return outerNodeEnd; }

    // helper method to resolve nodeName from source
    private static String resolveNodeName(XmlDocument sourceDocument, int currentNodeIndex, int nextNodeIndex) throws Exception {

        int _firstSpace = sourceDocument.indexOfBetween(" ", currentNodeIndex, nextNodeIndex );
        int _firstGt = sourceDocument.indexOfBetween(">", currentNodeIndex, nextNodeIndex );

        if (_firstSpace >= 0 && _firstGt < 0)
            return sourceDocument.xmlString.substring(currentNodeIndex + 1, _firstSpace);
        else if (_firstSpace < 0 && _firstGt >= 0)
            return sourceDocument.xmlString.substring(currentNodeIndex + 1, _firstGt);
        else if (_firstSpace >= 0 && _firstGt >= 0)
            return sourceDocument.xmlString.substring(currentNodeIndex + 1, Math.min(_firstSpace, _firstGt));
        else
            throw new Exception("XML is not valid!");
    }

    // helper method to resolve nodeName
    public static String resolveNodeNameFromXmlString(String xmlNode) throws Exception {

        if (!xmlNode.startsWith("<"))
            throw new Exception("XML is not valid!");

        int _firstSpace = xmlNode.indexOf(" ");
        int _firstGt = xmlNode.indexOf(">");

        if (_firstSpace >= 0 && _firstGt < 0)
            return xmlNode.substring(1, _firstSpace);
        else if (_firstSpace < 0 && _firstGt >= 0)
            return xmlNode.substring(1, _firstGt);
        else if (_firstSpace >= 0 && _firstGt >= 0)
            return xmlNode.substring(1, Math.min(_firstSpace, _firstGt));
        else
            throw new Exception("XML is not valid!");
    }
}
package com.company.customXmlParser;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class NodeSdt extends NodeDefault {

    private boolean nodeIsFullyInitialized = false;

    private String theNodeTemplate = null;
    public String getTheNodeTemplate()
    {
        return theNodeTemplate;
    }

    private String theNodeName = null;

    private String bindingProperty = "";
    public String getBindingPropertyName()
    {
        return bindingProperty;
    }

    // constructor accessible only inside the package
    NodeSdt(){

    }

    @Override
    public void setOuterNodeEnd(int idx){

        super.setOuterNodeEnd(idx);

        try {
            this.initializeTheNodeState();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public String BindTheDataAndGetXml(String textData)
    {
        if (!nodeIsFullyInitialized)
            return super.getOuterXml();

        String[] TextLines;
        if (theNodeName.equals("w:p"))
            TextLines = textData.split(Pattern.quote("\n"));
        else
            TextLines = new String[] {textData};

        StringBuilder toReturn = new StringBuilder();
        for(String toInsert: TextLines)
        {
            int L;
            if ((L = theNodeTemplate.indexOf("<w:t ")) > 0 || (L = theNodeTemplate.indexOf("<w:t>")) > 0)
            {
                L = theNodeTemplate.indexOf(">", L + 1);
                if (L > 0)
                {
                    int R = theNodeTemplate.indexOf("</w:t>", L + 1);
                    if (R > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.append(theNodeTemplate.substring(0, L + 1));
                        toInsert = org.apache.commons.lang3.StringEscapeUtils.escapeXml10(toInsert);
                        sb.append(toInsert);
                        sb.append(theNodeTemplate.substring(R));
                        toReturn.append(sb.toString());
                    }
                }
            }
        }

        return toReturn.toString();
    }

    private void initializeTheNodeState() throws Exception {

        String originalOuterXml = super.getOuterXml();

        // theNodeTemplate
        String regexString = Pattern.quote("<w:sdtContent>") + "(.*?)" + Pattern.quote("</w:sdtContent>");
        Pattern pattern = Pattern.compile(regexString);
        Matcher matcher = pattern.matcher(originalOuterXml);
        if (matcher.find()) {
            theNodeTemplate = matcher.group(1);
        }

        // theNodeName
        theNodeName =  super.resolveNodeNameFromXmlString(theNodeTemplate);

        // bindingProperty
        regexString = Pattern.quote("<w:tag ") + "(.*?)" + Pattern.quote("/>");
        pattern = Pattern.compile(regexString);
        matcher = pattern.matcher(originalOuterXml);
        if (matcher.find()) {
            matcher = Pattern.compile(Pattern.quote("\"") + "(.*?)" + Pattern.quote("\"")).matcher(matcher.group(1));
            if (matcher.find())
                bindingProperty = matcher.group(1);
        }

        nodeIsFullyInitialized = true;
    }
}
package com.company.customXmlParser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class NodeTbl extends NodeDefault {

    private boolean nodeIsFullyInitialized = false;

    private String theNodeTemplate = null;
    public String getTheNodeTemplate()
    {
        return theNodeTemplate;
    }

  //  private String theNodeName = null;

    private String bindingProperty = "";
    public String getBindingPropertyName()
    {
        return bindingProperty;
    }

    private ArrayList<int[]> arrayOfRows = null;

    // constructor accessible only inside the package
    NodeTbl(){

    }

    @Override
    public void setOuterNodeEnd(int idx){

        super.setOuterNodeEnd(idx);

        try {
            this.initializeTheNodeState();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public String BindTheDataAndGetXml(Object data)
    {
        if (data instanceof ArrayList) {

            StringBuilder toReturn = new StringBuilder();
            try
            {
                String xmlHeaderRow = theNodeTemplate.substring(arrayOfRows.get(0)[0], arrayOfRows.get(0)[1]);
                xmlHeaderRow = (new XmlDocument(xmlHeaderRow)).BindTheDataAndGetXml(new HashMap(), IfDataNodeIsNull.RemoveTheWholeNode);

                String xmlRowTemplate = theNodeTemplate.substring(arrayOfRows.get(1)[0], arrayOfRows.get(1)[1]);

                StringBuilder resultRows = new StringBuilder(xmlHeaderRow);

                ArrayList<HashMap<String, Object>> dataArray = (ArrayList<HashMap<String, Object>>) data;
                for (int i = 0; i < dataArray.size(); i++) {
                    HashMap<String, Object> dataToInsert = dataArray.get(i);
                    XmlDocument xmlRow = new XmlDocument(xmlRowTemplate);
                    resultRows.append(xmlRow.BindTheDataAndGetXml(dataToInsert, IfDataNodeIsNull.DontChange));
                }

                toReturn.append(theNodeTemplate.substring(0, arrayOfRows.get(0)[0]));
                toReturn.append(resultRows);
                toReturn.append(theNodeTemplate.substring(arrayOfRows.get(1)[1]));

            } catch (Exception e) {
                e.printStackTrace();
            }

            return toReturn.toString();
        }
        else
            return theNodeTemplate;
    }

    private void initializeTheNodeState() throws Exception {

        //
        String originalOuterXml = super.getOuterXml();

        // theNodeTemplate
        theNodeTemplate = originalOuterXml;
        fillTheArrayOfRows();

      //  // newNodeName
      //  theNodeName =  resolveNodeNameFromXmlString(theNodeTemplate);

        // bindingProperty
        String regexString = Pattern.quote("<w:tag ") + "(.*?)" + Pattern.quote("/>");
        Pattern pattern = Pattern.compile(regexString);
        Matcher matcher = pattern.matcher(originalOuterXml);
        if (matcher.find()) {
            matcher = Pattern.compile(Pattern.quote("\"") + "(.*?)" + Pattern.quote("\"")).matcher(matcher.group(1));
            if (matcher.find()) {
                String fullTagName = matcher.group(1);
                if (fullTagName.startsWith("CiProTable_")){
                    bindingProperty = fullTagName.substring("CiProTable_".length());
                    nodeIsFullyInitialized = true;
                }
            }
        }
    }

    private void fillTheArrayOfRows()
    {
        arrayOfRows = new ArrayList<>();

        int L = 0;
        while ((L = theNodeTemplate.indexOf("<w:tr ", L + 1)) >= 0)
        {
            int R;
            if ((R = theNodeTemplate.indexOf("</w:tr>", L + 1)) > 0)
                arrayOfRows.add(new int[]{L, R + "</w:tr>".length()});
            else
                break;
        }
    }
}
package com.company.customXmlParser;

import java.util.ArrayList;
import java.util.HashMap;

public class NodeVirtualBlock extends NodeDefault  {

    private boolean nodeIsFullyInitialized = false;

    private String theNodeTemplate = null;
    public String getTheNodeTemplate()
    {
        return theNodeTemplate;
    }

    private String theNodeName = null;

    private String bindingProperty = "";
    public String getBindingPropertyName()
    {
        return bindingProperty;
    }

    // constructor accessible only inside the package
    NodeVirtualBlock(){

    }

    @Override
    public void setOuterNodeEnd(int idx){

        super.setOuterNodeEnd(idx);

        try {
            this.initializeTheNodeState();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public String BindTheDataAndGetXml(Object data)
    {
        if (data instanceof ArrayList) {

            StringBuilder toReturn = new StringBuilder();
            try {

                ArrayList<HashMap<String, Object>> dataArray = (ArrayList<HashMap<String, Object>>) data;
                for (int i = 0; i < dataArray.size(); i++) {
                    HashMap<String, Object> dataToInsert = dataArray.get(i);

                    if (children.size() > 3)
                    for(int k = 1; k < children.size() - 1; k++) {

                        XmlDocument xmlBlock = new XmlDocument(children.get(k).getOuterXml());
                        toReturn.append(xmlBlock.BindTheDataAndGetXml(dataToInsert, IfDataNodeIsNull.DontChange));

                    }
                }

            } catch (Exception e) {
                e.printStackTrace();
            }

            return toReturn.toString();
        }
        else
            return theNodeTemplate;

    }

    private void initializeTheNodeState() throws Exception {

        // newNodeTemplate
        theNodeTemplate = "" ;
        //fillTheArrayOfRows();

        // newNodeName
        theNodeName = "";

        // bindingProperty
        bindingProperty = ((NodeSdt)children.get(0)).getBindingPropertyName().substring("CiProLoopBegin_".length());
        nodeIsFullyInitialized = true;

    }


}
