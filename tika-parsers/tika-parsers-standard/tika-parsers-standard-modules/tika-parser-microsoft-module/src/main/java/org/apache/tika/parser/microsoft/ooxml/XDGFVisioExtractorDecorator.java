/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.apache.tika.parser.microsoft.ooxml;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ooxml.extractor.POIXMLTextExtractor;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xdgf.usermodel.XDGFConnection;
import org.apache.poi.xdgf.usermodel.XDGFPage;
import org.apache.poi.xdgf.usermodel.XDGFShape;
import org.apache.poi.xdgf.usermodel.XmlVisioDocument;
import org.xml.sax.SAXException;

import org.apache.tika.parser.ParseContext;
import org.apache.tika.sax.XHTMLContentHandler;

/**
 * Extracts content from OOXML Visio (.vsdx) documents and emits a structured
 * Mermaid flowchart representation preserving shapes and connections.
 */
public class XDGFVisioExtractorDecorator extends AbstractOOXMLExtractor {

    public XDGFVisioExtractorDecorator(ParseContext context, POIXMLTextExtractor extractor) {
        super(context, extractor);
    }

    @Override
    protected void buildXHTML(XHTMLContentHandler xhtml) throws SAXException {
        XmlVisioDocument document = (XmlVisioDocument) extractor.getDocument();

        for (XDGFPage page : document.getPages()) {
            xhtml.startElement("div", "class", "page");
            xhtml.element("h1", page.getName());

            Map<Long, String> shapeNodeIds = new LinkedHashMap<>();
            Map<Long, String> shapeLabels = new LinkedHashMap<>();
            List<String> edges = new ArrayList<>();

            collectShapes(page.getContent().getTopLevelShapes(), shapeNodeIds, shapeLabels);
            collectEdges(page.getContent().getConnections(), shapeNodeIds, shapeLabels, edges);

            StringBuilder mermaid = new StringBuilder();
            mermaid.append("flowchart TD\n");

            for (Map.Entry<Long, String> entry : shapeNodeIds.entrySet()) {
                String nodeId = entry.getValue();
                String label = shapeLabels.get(entry.getKey());
                if (label != null && !label.isEmpty()) {
                    mermaid.append("    ").append(nodeId)
                            .append("[\"").append(escapeMermaid(label)).append("\"]\n");
                }
            }

            for (String edge : edges) {
                mermaid.append("    ").append(edge).append("\n");
            }

            xhtml.startElement("pre", "class", "mermaid");
            xhtml.characters(mermaid.toString());
            xhtml.endElement("pre");

            xhtml.endElement("div");
        }
    }

    private void collectShapes(List<XDGFShape> shapes, Map<Long, String> shapeNodeIds,
                               Map<Long, String> shapeLabels) {
        if (shapes == null) {
            return;
        }
        for (XDGFShape shape : shapes) {
            if (shape.isDeleted()) {
                continue;
            }
            // 1D shapes are connectors — skip them as nodes
            if (shape.isShape1D()) {
                continue;
            }

            long id = shape.getID();
            String nodeId = "s" + id;
            shapeNodeIds.put(id, nodeId);

            String text = "";
            if (shape.hasText()) {
                text = shape.getTextAsString().trim();
            }
            if (text.isEmpty() && shape.getName() != null) {
                text = shape.getName();
            }
            shapeLabels.put(id, text);

            // recurse into grouped shapes
            if (shape.hasShapes()) {
                collectShapes(shape.getShapes(), shapeNodeIds, shapeLabels);
            }
        }
    }

    private void collectEdges(List<XDGFConnection> connections, Map<Long, String> shapeNodeIds,
                              Map<Long, String> shapeLabels, List<String> edges) {
        if (connections == null) {
            return;
        }

        // De-duplicate connections: a connector often has two Connect elements,
        // one for each end. We group by the 1D connector shape (fromShape) and
        // collect the target shapes it connects.
        Map<Long, List<XDGFShape>> connectorTargets = new LinkedHashMap<>();

        for (XDGFConnection conn : connections) {
            XDGFShape from = conn.getFromShape();
            XDGFShape to = conn.getToShape();
            if (from == null || to == null) {
                continue;
            }

            // The "from" shape in a Connect element is the connector (1D shape).
            // The "to" shape is the shape it glues to.
            if (from.isShape1D()) {
                connectorTargets.computeIfAbsent(from.getID(), k -> new ArrayList<>()).add(to);
            } else if (to.isShape1D()) {
                connectorTargets.computeIfAbsent(to.getID(), k -> new ArrayList<>()).add(from);
            } else {
                // Direct shape-to-shape connection (rare) — emit directly
                ensureNode(from, shapeNodeIds, shapeLabels);
                ensureNode(to, shapeNodeIds, shapeLabels);
                edges.add(shapeNodeIds.get(from.getID()) + " --> " +
                        shapeNodeIds.get(to.getID()));
            }
        }

        // For each connector that touches exactly 2 shapes, emit an edge
        for (Map.Entry<Long, List<XDGFShape>> entry : connectorTargets.entrySet()) {
            List<XDGFShape> targets = entry.getValue();
            if (targets.size() == 2) {
                XDGFShape a = targets.get(0);
                XDGFShape b = targets.get(1);
                ensureNode(a, shapeNodeIds, shapeLabels);
                ensureNode(b, shapeNodeIds, shapeLabels);

                // Try to get label from the connector shape itself
                String connectorLabel = getConnectorLabel(
                        ((XmlVisioDocument) extractor.getDocument()), entry.getKey());
                String arrow = " --> ";
                if (connectorLabel != null && !connectorLabel.isEmpty()) {
                    arrow = " -->|\"" + escapeMermaid(connectorLabel) + "\"| ";
                }
                edges.add(shapeNodeIds.get(a.getID()) + arrow +
                        shapeNodeIds.get(b.getID()));
            } else if (targets.size() == 1) {
                // Connector with only one end connected — emit a dangling node
                XDGFShape a = targets.get(0);
                ensureNode(a, shapeNodeIds, shapeLabels);
            }
        }
    }

    private void ensureNode(XDGFShape shape, Map<Long, String> shapeNodeIds,
                            Map<Long, String> shapeLabels) {
        long id = shape.getID();
        if (!shapeNodeIds.containsKey(id)) {
            shapeNodeIds.put(id, "s" + id);
            String text = "";
            if (shape.hasText()) {
                text = shape.getTextAsString().trim();
            }
            if (text.isEmpty() && shape.getName() != null) {
                text = shape.getName();
            }
            shapeLabels.put(id, text);
        }
    }

    private String getConnectorLabel(XmlVisioDocument doc, long connectorShapeId) {
        for (XDGFPage page : doc.getPages()) {
            XDGFShape shape = page.getContent().getShapeById(connectorShapeId);
            if (shape != null && shape.hasText()) {
                return shape.getTextAsString().trim();
            }
        }
        return null;
    }

    private String escapeMermaid(String text) {
        return text.replace("\\", "\\\\")
                .replace("\"", "#quot;")
                .replace("\n", " ")
                .replace("\r", "");
    }

    @Override
    protected List<PackagePart> getMainDocumentParts() {
        return new ArrayList<>();
    }
}
