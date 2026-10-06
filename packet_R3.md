# 评分包 R3

共 60 条。每条只需判定该 source 适用的维度，判定规则见 `rater_guide.md`。
把答案填进 `packet_R3.csv`，一行一条；拿不准就填 `unclear` 并在 notes 里写一句原因。

---

## IRR-001  ·  MethodSource

**项目** `Commons-RDF`  **文件** `commons-rdf/commons-rdf-integration-tests/src/test/java/org/apache/commons/rdf/integrationtests/AllToAllTest.java`  **测试** `testAddTermsFromOtherFactory`

### 测试方法

```java
    @MethodSource("data")
    @ParameterizedTest(name = "{index}: {0} -> {1}")
    void testAddTermsFromOtherFactory(final Class<? extends RDF> from, final Class<? extends RDF> to) throws Exception {
        RDF nodeFactory = from.getConstructor().newInstance();
        RDF graphFactory = to.newInstance();

        try (final Graph g = graphFactory.createGraph()) {
            final BlankNode s = nodeFactory.createBlankNode();
            final IRI p = nodeFactory.createIRI("http://example.com/p");
            final Literal o = nodeFactory.createLiteral("Hello");

            g.add(s, p, o);

            // blankNode should still work with g.contains()
            assertTrue(g.contains(s, p, o));
            final Triple t1 = g.stream().findAny().get();

            // Can't make assumptions about BlankNode equality - it might
            // have been mapped to a different BlankNode.uniqueReference()
            // assertEquals(s, t.getSubject());

            assertEquals(p, t1.getPredicate());
            assertEquals(o, t1.getObject());

            final IRI s2 = nodeFactory.createIRI("http://example.com/s2");
            g.add(s2, p, s);
            assertTrue(g.contains(s2, p, s));

            // This should be mapped to the same BlankNode
            // (even if it has a different identifier), e.g.
            // we should be able to do:

            final Triple t2 = g.stream(s2, p, null).findAny().get();

            final BlankNode bnode = (BlankNode) t2.getObject();
            // And that (possibly adapted) BlankNode object should
            // match the subject of t1 statement
            assertEquals(bnode, t1.getSubject());
            // And can be used as a key:
            final Triple t3 = g.stream(bnode, p, null).findAny().get();
            assertEquals(t1, t3);
        }
    }
```

### Parameter provider — 同文件内的 `data`

```java
    @SuppressWarnings("rawtypes")
    public static Collection<Object[]> data() {
        final List<Class> factories = Arrays.asList(SimpleRDF.class, JenaRDF.class, RDF4J.class, JsonLdRDF.class);
        final Collection<Object[]> allToAll = new ArrayList<>();
        for (final Class from : factories) {
            for (final Class to : factories) {
                // NOTE: we deliberately include self-to-self here
                // to test two instances of the same implementation
                allToAll.add(new Object[] { from, to });
            }
        }
        return allToAll;
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-002  ·  MethodSource

**项目** `POI`  **文件** `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/TestFormulaEvaluatorOnXSSF.java`  **测试** `processFunctionRow`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource("data")
    void processFunctionRow(String targetFunctionName, int formulasRowIdx, int expectedValuesRowIdx) {
        //DOLLAR function returns a string that is locale specific
        assumeFalse(targetFunctionName.equalsIgnoreCase("DOLLAR"));

        Row formulasRow = sheet.getRow(formulasRowIdx);
        Row expectedValuesRow = sheet.getRow(expectedValuesRowIdx);

        short endcolnum = formulasRow.getLastCellNum();

        // iterate across the row for all the evaluation cases
        for (short colnum=SS.COLUMN_INDEX_FIRST_TEST_VALUE; colnum < endcolnum; colnum++) {
            Cell c = formulasRow.getCell(colnum);
            assumeTrue(c != null);
            assumeTrue(c.getCellType() == CellType.FORMULA);
            ignoredFormulaTestCase(c.getCellFormula());

            CellValue actValue = evaluator.evaluate(c);
            Cell expValue = (expectedValuesRow == null) ? null : expectedValuesRow.getCell(colnum);

            String msg = String.format(Locale.ROOT, "Function '%s': Formula: %s @ %d:%d"
                , targetFunctionName, c.getCellFormula(), formulasRow.getRowNum(), colnum);

            assertNotNull(expValue, msg + " - Bad setup data expected value is null");
            assertNotNull(actValue, msg + " - actual value was null");

            final CellType expectedCellType = expValue.getCellType();
            switch (expectedCellType) {
                case BLANK:
                    assertEquals(CellType.BLANK, actValue.getCellType(), msg);
                    break;
                case BOOLEAN:
                    assertEquals(CellType.BOOLEAN, actValue.getCellType(), msg);
                    assertEquals(expValue.getBooleanCellValue(), actValue.getBooleanValue(), msg);
                    break;
                case ERROR:
                    assertEquals(CellType.ERROR, actValue.getCellType(), msg);
//                if(false) { // TODO: fix ~45 functions which are currently returning incorrect error values
//                    assertEquals(msg, expValue.getErrorCellValue(), actValue.getErrorValue());
//                }
                    break;
                case FORMULA: // will never be used, since we will call method after formula evaluation
                    fail("Cannot expect formula as result of formula evaluation: " + msg);
                case NUMERIC:
                    assertEquals(CellType.NUMERIC, actValue.getCellType(), msg);
                    final double tolerance = targetFunctionName.equalsIgnoreCase("RATE")
                            ? 0.000001 : BaseTestNumeric.DIFF_TOLERANCE_FACTOR;
                    BaseTestNumeric.assertDouble(msg, expValue.getNumericCellValue(), actValue.getNumberValue(), BaseTestNumeric.POS_ZERO, tolerance);
                    break;
                case STRING:
                    assertEquals(CellType.STRING, actValue.getCellType(), msg);
                    assertEquals(expValue.getRichStringCellValue().getString(), actValue.getStringValue(), msg);
                    break;
                default:
                    fail("Unexpected cell type: " + expectedCellType);
            }
        }
    }
```

### Parameter provider — 同文件内的 `data`

```java

    public static Stream<Arguments> data() throws Exception {
        // Function "Text" uses custom-formats which are locale specific
        // can't set the locale on a per-testrun execution, as some settings have been
        // already set, when we would try to change the locale by then
        userLocale = LocaleUtil.getUserLocale();
        LocaleUtil.setUserLocale(Locale.ROOT);

        workbook = new XSSFWorkbook( OPCPackage.open(HSSFTestDataSamples.getSampleFile(SS.FILENAME), PackageAccess.READ) );
        sheet = workbook.getSheetAt( 0 );
        evaluator = new XSSFFormulaEvaluator(workbook);

        List<Arguments> data = new ArrayList<>();

        processFunctionGroup(data, SS.START_OPERATORS_ROW_INDEX, null);
        processFunctionGroup(data, SS.START_FUNCTIONS_ROW_INDEX, null);
        // example for debugging individual functions/operators:
        // processFunctionGroup(data, SS.START_OPERATORS_ROW_INDEX, "ConcatEval");
        // processFunctionGroup(data, SS.START_FUNCTIONS_ROW_INDEX, "Text");

        return data.stream();
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-003  ·  EnumSource

**项目** `jena`  **文件** `jena/jena-ontapi/src/test/java/org/apache/jena/ontapi/OntClassIndividualsTest.java`  **测试** `testListIndividuals7a`

### 测试方法

```java
    public void testListIndividuals7a(TestSpec spec) {
        //  A   B
        //  .\ /.
        //  . C .
        //  . | .
        //  . D .
        //  ./  .
        //  A   .   E
        //   \  .  |
        //    \ . /
        //      B

        OntModel m = createClassesABCDAEB(OntModelFactory.createModel(spec.inst));
        OntClass A = m.getResource(NS + "A").as(OntClass.class);
        OntClass B = m.getResource(NS + "B").as(OntClass.class);
        OntClass C = m.getResource(NS + "C").as(OntClass.class);
        m.getResource(NS + "D").as(OntClass.class);
        OntClass E = m.getResource(NS + "E").as(OntClass.class);

        A.createIndividual(NS + "iA");
        B.createIndividual(NS + "iB");
        OntIndividual CE = C.createIndividual(NS + "iCE");
        CE.attachClass(E);
        OntIndividual DBA = B.createIndividual(NS + "iDBA");
        DBA.attachClass(B);
        DBA.attachClass(A);

        Set<String> directA = individuals(m, "A", true);
        Set<String> indirectA = individuals(m, "A", false);

        Set<String> directB = individuals(m, "B", true);
        Set<String> indirectB = individuals(m, "B", false);

        Set<String> directC = individuals(m, "C", true);
        Set<String> indirectC = individuals(m, "C", false);

        Set<String> directD = individuals(m, "D", true);
        Set<String> indirectD = individuals(m, "D", false);

        Set<String> directE = individuals(m, "E", true);
        Set<String> indirectE = individuals(m, "E", false);

        Assertions.assertEquals(Set.of("iA"), directA);
        Assertions.assertEquals(Set.of("iB", "iDBA"), directB);
        Assertions.assertEquals(Set.of("iCE"), directC);
        Assertions.assertEquals(Set.of(), directD);
        Assertions.assertEquals(Set.of("iCE"), directE);
        Assertions.assertEquals(Set.of("iA", "iDBA"), indirectA);
        Assertions.assertEquals(Set.of("iB", "iDBA"), indirectB);
        Assertions.assertEquals(Set.of("iCE"), indirectC);
        Assertions.assertEquals(Set.of(), indirectD);
        Assertions.assertEquals(Set.of("iCE"), indirectE);
    }
```

### 枚举声明 — `TestSpec`（jena/jena-ontapi/src/test/java/org/apache/jena/ontapi/TestSpec.java）

```java
public enum TestSpec {
    OWL2_MEM(OntSpecification.OWL2_FULL_MEM),
    OWL2_MEM_RDFS_INF(OntSpecification.OWL2_FULL_MEM_RDFS_INF),
    OWL2_MEM_TRANS_INF(OntSpecification.OWL2_FULL_MEM_TRANS_INF),
    OWL2_MEM_RULES_INF(OntSpecification.OWL2_FULL_MEM_RULES_INF),
    OWL2_MEM_MINI_RULES_INF(OntSpecification.OWL2_FULL_MEM_MINI_RULES_INF),
    OWL2_MEM_MICRO_RULES_INF(OntSpecification.OWL2_FULL_MEM_MICRO_RULES_INF),

    OWL2_DL_MEM_RDFS_BUILTIN_INF(OntSpecification.OWL2_DL_MEM_BUILTIN_RDFS_INF),
    OWL2_DL_MEM(OntSpecification.OWL2_DL_MEM),
    OWL2_DL_MEM_RDFS_INF(OntSpecification.OWL2_DL_MEM_RDFS_INF),
    OWL2_DL_MEM_TRANS_INF(OntSpecification.OWL2_DL_MEM_TRANS_INF),
    OWL2_DL_MEM_RULES_INF(OntSpecification.OWL2_DL_MEM_RULES_INF),

    OWL2_EL_MEM(OntSpecification.OWL2_EL_MEM),
    OWL2_EL_MEM_RDFS_INF(OntSpecification.OWL2_EL_MEM_RDFS_INF),
    OWL2_EL_MEM_TRANS_INF(OntSpecification.OWL2_EL_MEM_TRANS_INF),
    OWL2_EL_MEM_RULES_INF(OntSpecification.OWL2_EL_MEM_RULES_INF),

    OWL2_QL_MEM(OntSpecification.OWL2_QL_MEM),
    OWL2_QL_MEM_RDFS_INF(OntSpecification.OWL2_QL_MEM_RDFS_INF),
    OWL2_QL_MEM_TRANS_INF(OntSpecification.OWL2_QL_MEM_TRANS_INF),
    OWL2_QL_MEM_RULES_INF(OntSpecification.OWL2_QL_MEM_RULES_INF),

    OWL2_RL_MEM(OntSpecification.OWL2_RL_MEM),
    OWL2_RL_MEM_RDFS_INF(OntSpecification.OWL2_RL_MEM_RDFS_INF),
    OWL2_RL_MEM_TRANS_INF(OntSpecification.OWL2_RL_MEM_TRANS_INF),
    OWL2_RL_MEM_RULES_INF(OntSpecification.OWL2_RL_MEM_RULES_INF),

    OWL1_MEM(OntSpecification.OWL1_FULL_MEM),
    OWL1_MEM_RDFS_INF(OntSpecification.OWL1_FULL_MEM_RDFS_INF),
    OWL1_MEM_TRANS_INF(OntSpecification.OWL1_FULL_MEM_TRANS_INF),
    OWL1_MEM_RULES_INF(OntSpecification.OWL1_FULL_MEM_RULES_INF),
    OWL1_MEM_MINI_RULES_INF(OntSpecification.OWL1_FULL_MEM_MINI_RULES_INF),
    OWL1_MEM_MICRO_RULES_INF(OntSpecification.OWL1_FULL_MEM_MICRO_RULES_INF),

    OWL1_DL_MEM(OntSpecification.OWL1_DL_MEM),
    OWL1_DL_MEM_RDFS_INF(OntSpecification.OWL1_DL_MEM_RDFS_INF),
    OWL1_DL_MEM_TRANS_INF(OntSpecification.OWL1_DL_MEM_TRANS_INF),
    OWL1_DL_MEM_RULES_INF(OntSpecification.OWL1_DL_MEM_RULES_INF),

    OWL1_LITE_MEM(OntSpecification.OWL1_LITE_MEM),
    OWL1_LITE_MEM_RDFS_INF(OntSpecification.OWL1_LITE_MEM_RDFS_INF),
    OWL1_LITE_MEM_TRANS_INF(OntSpecification.OWL1_LITE_MEM_TRANS_INF),
    OWL1_LITE_MEM_RULES_INF(OntSpecification.OWL1_LITE_MEM_RULES_INF),

    RDFS_MEM(OntSpecification.RDFS_MEM),
    RDFS_MEM_RDFS_INF(OntSpecification.RDFS_MEM_RDFS_INF),
    RDFS_MEM_TRANS_INF(OntSpecification.RDFS_MEM_TRANS_INF),
    ;
    public final OntSpecification inst;

    TestSpec(OntSpecification inst) {
        this.inst = inst;
    }

    boolean isOWL1() {
        return name().startsWith("OWL1");
    }

    boolean isOWL1Lite() {
        return name().startsWith("OWL1_LITE");
    }

    boolean isOWL2() {
        return name().startsWith("OWL2");
    }

    boolean isOWL2EL() {
        return name().startsWith("OWL2_EL");
    }

    boolean isOWL2QL() {
        return name().startsWith("OWL2_QL");
    }

    boolean isOWL2RL() {
        return name().startsWith("OWL2_RL");
    }

    boolean isRules() {
        return name().endsWith("_RULES_INF");
    }

    boolean isRDFS() {
        return name().endsWith("_RDFS_INF");
    }
}
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-004  ·  EnumSource

**项目** `Hudi`  **文件** `hudi/hudi-io/src/test/java/org/apache/hudi/io/compress/TestHoodieCompressor.java`  **测试** `testDefaultDecompressors`

### 测试方法

```java
  @ParameterizedTest
  @EnumSource(CompressionCodec.class)
  public void testDefaultDecompressors(CompressionCodec codec) throws IOException {
    switch (codec) {
      case NONE:
      case GZIP:
        HoodieCompressor decompressor = HoodieCompressorFactory.getCompressor(codec);
        byte[] actualOutput = new byte[INPUT_LENGTH + 100];
        try (InputStream stream = prepareInputStream(codec)) {
          for (int sizeToRead : READ_PART_SIZE_LIST) {
            stream.mark(INPUT_LENGTH);
            int actualSizeRead =
                decompressor.decompress(stream, actualOutput, 4, sizeToRead);
            assertEquals(actualSizeRead, Math.min(INPUT_LENGTH, sizeToRead));
            assertEquals(0, IOUtils.compareTo(
                actualOutput, 4, actualSizeRead, INPUT_BYTES, 0, actualSizeRead));
            stream.reset();
          }
        }
        break;
      default:
        assertThrows(
            IllegalArgumentException.class, () -> HoodieCompressorFactory.getCompressor(codec));
    }
  }
```

### 枚举声明 — `CompressionCodec`（hudi/hudi-io/src/main/java/org/apache/hudi/io/compress/CompressionCodec.java）

```java
public enum CompressionCodec {
  NONE("none", 2),
  BZIP2("bz2", 5),
  GZIP("gz", 1),
  LZ4("lz4", 4),
  LZO("lzo", 0),
  SNAPPY("snappy", 3),
  ZSTD("zstd", 6);

  private static final Map<String, CompressionCodec>
      NAME_TO_COMPRESSION_CODEC_MAP = createNameToCompressionCodecMap();
  private static final Map<Integer, CompressionCodec>
      ID_TO_COMPRESSION_CODEC_MAP = createIdToCompressionCodecMap();

  private final String name;
  // CompressionCodec ID to be stored in HFile on storage
  // The ID of each codec cannot change or else that breaks all existing HFiles out there
  // even the ones that are not compressed! (They use the NONE algorithm)
  private final int id;

  CompressionCodec(final String name, int id) {
    this.name = name;
    this.id = id;
  }

  public String getName() {
    return name;
  }

  public int getId() {
    return id;
  }

  public static CompressionCodec findCodecByName(String name) {
    CompressionCodec codec =
        NAME_TO_COMPRESSION_CODEC_MAP.get(name.toLowerCase());
    ValidationUtils.checkArgument(
        codec != null, String.format("Cannot find compression codec: %s", name));
    return codec;
  }

  /**
   * Gets the compression codec based on the ID.  This ID is written to the HFile on storage.
   *
   * @param id ID indicating the compression codec
   * @return compression codec based on the ID
   */
  public static CompressionCodec decodeCompressionCodec(int id) {
    CompressionCodec codec = ID_TO_COMPRESSION_CODEC_MAP.get(id);
    ValidationUtils.checkArgument(
        codec != null, "Compression code not found for ID: " + id);
    return codec;
  }

  /**
   * @return the mapping of name to compression codec.
   */
  private static Map<String, CompressionCodec> createNameToCompressionCodecMap() {
    return Collections.unmodifiableMap(
        Arrays.stream(CompressionCodec.values())
            .collect(Collectors.toMap(CompressionCodec::getName, Function.identity()))
    );
  }

  /**
   * @return the mapping of ID to compression codec.
   */
  private static Map<Integer, CompressionCodec> createIdToCompressionCodecMap() {
    return Collections.unmodifiableMap(
        Arrays.stream(CompressionCodec.values())
            .collect(Collectors.toMap(CompressionCodec::getId, Function.identity()))
    );
  }
}
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-005  ·  ValueSource

**项目** `Maven`  **文件** `maven/impl/maven-core/src/test/java/org/apache/maven/plugin/PluginParameterExpressionEvaluatorTest.java`  **测试** `testValueExtractionOfMissingPrefixedSuffixedProperty`

### 测试方法

```java
    void testValueExtractionOfMissingPrefixedSuffixedProperty(String missingPropertyExpression) throws Exception {
        Properties executionProperties = new Properties();

        ExpressionEvaluator ee = createExpressionEvaluator(null, null, executionProperties);

        Object value = ee.evaluate(missingPropertyExpression);

        assertEquals(missingPropertyExpression, value);
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-006  ·  ValueSource

**项目** `commons-rng`  **文件** `commons-rng/commons-rng-sampling/src/test/java/org/apache/commons/rng/sampling/ArraySamplerTest.java`  **测试** `testShuffleIsRandom`

### 测试方法

```java
    @ParameterizedTest
    @ValueSource(ints = {13, 16})
    void testShuffleIsRandom(int length) {
        final int[] array = PermutationSampler.natural(length);
        final UniformRandomProvider rng = RandomAssert.createRNG();
        final long[][] counts = new long[length][length];
        for (int j = 1; j <= 1000; j++) {
            ArraySampler.shuffle(rng, array);
            for (int i = 0; i < length; i++) {
                counts[i][array[i]]++;
            }
        }
        final double p = new ChiSquareTest().chiSquareTest(counts);
        Assertions.assertFalse(p < 1e-3, () -> "p-value too small: " + p);
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-007  ·  MethodSource

**项目** `Commons-Compress`  **文件** `commons-compress/src/test/java/org/apache/commons/compress/changes/ChangeSetSafeTypesTest.java`  **测试** `testDeleteFileCpio`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource("org.apache.commons.compress.changes.TestFixtures#getOutputArchiveNames")
    void testDeleteFileCpio(final String archiverName) throws Exception {
        final Path input = createArchive(archiverName);
        final File result = createTempFile("test", "." + archiverName);
        try (InputStream inputStream = Files.newInputStream(input);
                ArchiveInputStream<E> ais = createArchiveInputStream(archiverName, inputStream);
                OutputStream outputStream = Files.newOutputStream(result.toPath());
                ArchiveOutputStream<E> out = createArchiveOutputStream(archiverName, outputStream)) {
            final ChangeSet<E> changeSet = createChangeSet();
            changeSet.delete("bla/test5.xml");
            archiveListDelete("bla/test5.xml");
            new ChangeSetPerformer<>(changeSet).perform(ais, out);
        }
        checkArchiveContent(result, archiveList);
    }
```

### Parameter provider — `TestFixtures#getOutputArchiveNames`（commons-compress/src/test/java/org/apache/commons/compress/changes/TestFixtures.java）

```java

    static Set<String> getOutputArchiveNames() {
        final Set<String> outputStreamArchiveNames = ArchiveStreamFactory.DEFAULT.getOutputStreamArchiveNames();
        outputStreamArchiveNames.remove(ArchiveStreamFactory.AR); // TODO BUG?
        outputStreamArchiveNames.remove(ArchiveStreamFactory.SEVEN_Z); // TODO Does not support streaming.
        return outputStreamArchiveNames;
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-008  ·  EnumSource

**项目** `camel`  **文件** `camel/components/camel-infinispan/camel-infinispan/src/test/java/org/apache/camel/component/infinispan/remote/InfinispanRemoteEmbeddingStoreIT.java`  **测试** `registerSchema`

### 测试方法

```java
    @ParameterizedTest
    @EnumSource(VectorSimilarity.class)
    public void registerSchema(VectorSimilarity similarity) {
        int dimension = 900 + similarity.ordinal();
        String typeName = EmbeddingStoreUtil.DEFAULT_TYPE_NAME_PREFIX + dimension;

        InfinispanRemoteConfiguration configuration = createInfinispanRemoteConfiguration();
        configuration.setEmbeddingStoreDimension(dimension);
        configuration.setEmbeddingStoreTypeName(typeName);
        configuration.setEmbeddingStoreVectorSimilarity(similarity);

        InfinispanRemoteManager manager = new InfinispanRemoteManager(context, configuration);
        BasicCache<Object, Object> metadataCache = null;
        try {
            manager.start();

            metadataCache = manager.getCache(ProtobufMetadataManagerConstants.PROTOBUF_METADATA_CACHE_NAME);
            Object metadata = metadataCache.get(EmbeddingStoreUtil.getSchemeFileName(configuration));
            assertNotNull(metadata);
        } finally {
            if (metadataCache != null) {
                metadataCache.remove(EmbeddingStoreUtil.getSchemeFileName(configuration));
            }
            manager.stop();
        }
    }
```

### 枚举声明

> ⚠️ 未能定位枚举声明。

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-009  ·  MethodSource

**项目** `Log4j`  **文件** `logging-log4j2/log4j-api-test/src/test/java/org/apache/logging/log4j/util/PropertySourceTokenizerTest.java`  **测试** `testTokenize`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource("data")
    void testTokenize(final String value, final List<CharSequence> expectedTokens) {
        final List<CharSequence> tokens = PropertySource.Util.tokenize(value);
        assertEquals(expectedTokens, tokens);
    }
```

### Parameter provider — 同文件内的 `data`

```java

    public static Object[][] data() {
        return new Object[][] {
            {"log4j.simple", Collections.singletonList("simple")},
            {"log4j_simple", Collections.singletonList("simple")},
            {"log4j-simple", Collections.singletonList("simple")},
            {"log4j/simple", Collections.singletonList("simple")},
            {"log4j2.simple", Collections.singletonList("simple")},
            {"Log4jSimple", Collections.singletonList("simple")},
            {"LOG4J_simple", Collections.singletonList("simple")},
            {"org.apache.logging.log4j.simple", Collections.singletonList("simple")},
            {"log4j.simpleProperty", Arrays.asList("simple", "property")},
            {"log4j.simple_property", Arrays.asList("simple", "property")},
            {"LOG4J_simple_property", Arrays.asList("simple", "property")},
            {"LOG4J_SIMPLE_PROPERTY", Arrays.asList("simple", "property")},
            {"log4j2-dashed-propertyName", Arrays.asList("dashed", "property", "name")},
            {"Log4jProperty_with.all-the/separators", Arrays.asList("property", "with", "all", "the", "separators")},
            {"org.apache.logging.log4j.config.property", Arrays.asList("config", "property")},
            // LOG4J2-3413
            {"level", Collections.emptyList()},
            {"user.home", Collections.emptyList()},
            {"CATALINA_BASE", Collections.emptyList()}
        };
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-010  ·  EnumSource

**项目** `Druid`  **文件** `druid/processing/src/test/java/org/apache/druid/query/metadata/SegmentMetadataQueryQueryToolChestTest.java`  **测试** `testProjectionsWithNull`

### 测试方法

```java
  @EnumSource(AggregatorMergeStrategy.class)
  @ParameterizedTest(name = "{index}: with AggregatorMergeStrategy {0}")
  public void testProjectionsWithNull(AggregatorMergeStrategy aggregatorMergeStrategy)
  {
    final SegmentAnalysis analysis1 = new SegmentAnalysis.Builder(TEST_SEGMENT_ID1)
        .projection("channel_sum", new AggregateProjectionMetadata(PROJECTION_CHANNEL_ADDED_HOURLY, 100))
        .build();
    final SegmentAnalysis analysis1NullProjection = new SegmentAnalysis.Builder(TEST_SEGMENT_ID1).build();
    final SegmentAnalysis analysis2 = new SegmentAnalysis.Builder(TEST_SEGMENT_ID2)
        .projection("channel_sum", new AggregateProjectionMetadata(PROJECTION_CHANNEL_ADDED_HOURLY, 200))
        .build();
    final SegmentAnalysis analysis2NullProjection = new SegmentAnalysis.Builder(TEST_SEGMENT_ID2).build();

    Assert.assertNull(mergeWithStrategy(analysis1NullProjection, analysis2, aggregatorMergeStrategy).getProjections());
    Assert.assertNull(mergeWithStrategy(analysis1, analysis2NullProjection, aggregatorMergeStrategy).getProjections());
    Assert.assertNull(
        mergeWithStrategy(analysis1NullProjection, analysis2NullProjection, aggregatorMergeStrategy).getProjections()
    );
  }
```

### 枚举声明 — `AggregatorMergeStrategy`（druid/processing/src/main/java/org/apache/druid/query/metadata/metadata/AggregatorMergeStrategy.java）

```java
public enum AggregatorMergeStrategy
{
  STRICT,
  LENIENT,
  EARLIEST,
  LATEST;

  @JsonValue
  @Override
  public String toString()
  {
    return StringUtils.toLowerCase(this.name());
  }

  @JsonCreator
  public static AggregatorMergeStrategy fromString(String name)
  {
    return valueOf(StringUtils.toUpperCase(name));
  }
}
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-011  ·  CsvSource

**项目** `Commons-Lang`  **文件** `commons-lang/src/test/java/org/apache/commons/lang3/math/FractionTest.java`  **测试** `testHashCodeNotEquals`

### 测试方法

```java
    void testHashCodeNotEquals(final int f1n, final int f1d, final int f2n, final int f2d) {
        assertNotEquals(Fraction.getFraction(f1n, f1d), Fraction.getFraction(f2n, f2d));
        assertNotEquals(Fraction.getFraction(f1n, f1d).hashCode(), Fraction.getFraction(f2n, f2d).hashCode());
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-012  ·  MethodSource

**项目** `Hadoop`  **文件** `hadoop/hadoop-hdfs-project/hadoop-hdfs/src/test/java/org/apache/hadoop/hdfs/server/datanode/checker/TestDatasetVolumeChecker.java`  **测试** `testInvalidConfigurationValues`

### 测试方法

```java
  @ParameterizedTest(name="{0}")
  @MethodSource("data")
  public void testInvalidConfigurationValues(VolumeCheckResult pExpectedVolumeHealth)
      throws Exception {
    initTestDatasetVolumeChecker(pExpectedVolumeHealth);
    HdfsConfiguration conf = new HdfsConfiguration();
    conf.setInt(DFS_DATANODE_DISK_CHECK_TIMEOUT_KEY, 0);
    intercept(HadoopIllegalArgumentException.class,
        "Invalid value configured for dfs.datanode.disk.check.timeout"
            + " - 0 (should be > 0)",
        () -> new DatasetVolumeChecker(conf, new FakeTimer()));
    conf.unset(DFS_DATANODE_DISK_CHECK_TIMEOUT_KEY);

    conf.setInt(DFS_DATANODE_DISK_CHECK_MIN_GAP_KEY, -1);
    intercept(HadoopIllegalArgumentException.class,
        "Invalid value configured for dfs.datanode.disk.check.min.gap"
            + " - -1 (should be >= 0)",
        () -> new DatasetVolumeChecker(conf, new FakeTimer()));
    conf.unset(DFS_DATANODE_DISK_CHECK_MIN_GAP_KEY);

    conf.setInt(DFS_DATANODE_DISK_CHECK_TIMEOUT_KEY, -1);
    intercept(HadoopIllegalArgumentException.class,
        "Invalid value configured for dfs.datanode.disk.check.timeout"
            + " - -1 (should be > 0)",
        () -> new DatasetVolumeChecker(conf, new FakeTimer()));
    conf.unset(DFS_DATANODE_DISK_CHECK_TIMEOUT_KEY);

    conf.setInt(DFS_DATANODE_FAILED_VOLUMES_TOLERATED_KEY, -2);
    intercept(HadoopIllegalArgumentException.class,
        "Invalid value configured for dfs.datanode.failed.volumes.tolerated"
            + " - -2 should be greater than or equal to -1",
        () -> new DatasetVolumeChecker(conf, new FakeTimer()));
  }
```

### Parameter provider — 同文件内的 `data`

```java
  public static Collection<Object[]> data() {
    List<Object[]> values = new ArrayList<>();
    for (VolumeCheckResult result : VolumeCheckResult.values()) {
      values.add(new Object[] {result});
    }
    values.add(new Object[] {null});
    return values;
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-013  ·  ValueSource

**项目** `Commons-BCEL`  **文件** `commons-bcel/src/test/java/org/apache/bcel/generic/EmptyVisitorTest.java`  **测试** `test`

### 测试方法

```java
    void test(final String className) throws ClassNotFoundException {
        // "java.io.Bits" is not in Java 21.
        assumeFalse(SystemUtils.isJavaVersionAtLeast(JavaVersion.JAVA_21) && className.equals("java.io.Bits"));
        final JavaClass javaClass = SyntheticRepository.getInstance().loadClass(className);
        for (final Method method : javaClass.getMethods()) {
            final Code code = method.getCode();
            if (code != null) {
                final InstructionList instructionList = new InstructionList(code.getCode());
                for (final InstructionHandle instructionHandle : instructionList) {
                    instructionHandle.accept(new EmptyVisitor() {
                        @Override
                        public void visitBREAKPOINT(final BREAKPOINT obj) {
                            fail(RESERVED_OPCODE);
                        }

                        @Override
                        public void visitIMPDEP1(final IMPDEP1 obj) {
                            fail(RESERVED_OPCODE);
                        }

                        @Override
                        public void visitIMPDEP2(final IMPDEP2 obj) {
                            fail(RESERVED_OPCODE);
                        }
                    });
                }
            }
        }
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-014  ·  CsvSource

**项目** `Commons-RNG`  **文件** `commons-rng/commons-rng-client-api/src/test/java/org/apache/commons/rng/UniformRandomProviderTest.java`  **测试** `testNextDoubleUniform`

### 测试方法

```java
    void testNextDoubleUniform(long seed, double origin, double bound) {
        Assertions.assertEquals((long) origin, origin, "origin");
        Assertions.assertEquals((long) bound, bound, "bound");
        final UniformRandomProvider rng = createRNG(seed);
        // Note casting as long will round towards zero.
        // If the upper bound is negative then this can create a domain error so use floor.
        final LongSupplier nextMethod = origin == 0 ?
                () -> (long) rng.nextDouble(bound) :
                () -> (long) Math.floor(rng.nextDouble(origin, bound));
        checkNextInRange("nextDouble", (long) origin, (long) bound, nextMethod);
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-015  ·  EnumSource

**项目** `Avro`  **文件** `avro/lang/java/avro/src/test/java/org/apache/avro/TestReadingWritingDataInEvolvedSchemas.java`  **测试** `floatWrittenWithUnionSchemaIsNotConvertedToLongSchema`

### 测试方法

```java
  @ParameterizedTest
  @EnumSource(EncoderType.class)
  void floatWrittenWithUnionSchemaIsNotConvertedToLongSchema(EncoderType encoderType) throws Exception {
    Schema writer = UNION_INT_LONG_FLOAT_DOUBLE_RECORD;
    Record record = defaultRecordWithSchema(writer, FIELD_A, 42.0f);
    byte[] encoded = encodeGenericBlob(record, encoderType);
    AvroTypeException exception = Assertions.assertThrows(AvroTypeException.class,
        () -> decodeGenericBlob(LONG_RECORD, writer, encoded, encoderType));
    Assertions.assertEquals("Found float, expecting long", exception.getMessage());
  }
```

### 枚举声明 — `EncoderType`（avro/lang/java/avro/src/test/java/org/apache/avro/TestReadingWritingDataInEvolvedSchemas.java）

```java
  enum EncoderType {
    BINARY, JSON
  }
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-018  ·  MethodSource

**项目** `uima-uimaj`  **文件** `uima-uimaj/uimaj-core/src/test/java/org/apache/uima/cas/serdes/CasSerializationDeserialization_XCAS_Test.java`  **测试** `roundTripDeserializeSerializeTest`

### 测试方法

```java
  @ParameterizedTest
  @MethodSource("roundTripDesSerScenarios")
  void roundTripDeserializeSerializeTest(Runnable aScenario) throws Exception {
    assumeNotKnownToFail(aScenario, //
            ".*casWithSofaDataArray",
            "XCAS does not suport SofA data arrays during deserialiaztion");

    aScenario.run();
  }
```

### Parameter provider — 同文件内的 `roundTripDesSerScenarios`

```java

  private static List<DesSerTestScenario> roundTripDesSerScenarios() throws Exception {
    return SerDesCasIOTestUtils.roundTripDesSerScenariosComparingCasContents(desSerCycles,
            CAS_FILE_NAME);
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-021  ·  MethodSource

**项目** `commons-numbers`  **文件** `commons-numbers/commons-numbers-examples/examples-jmh/src/test/java/org/apache/commons/numbers/examples/jmh/arrays/KthSelectorTest.java`  **测试** `testSelectSPN`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource(value = {"testSelect"})
    void testSelectSPN(double[] values) {
        final double[] sorted = values.clone();
        Arrays.sort(sorted);
        final KthSelector selector = new KthSelector();
        final double[] kp1 = new double[1];
        for (int i = 0; i < sorted.length; i++) {
            final int k = i;
            double[] x = values.clone();
            Assertions.assertEquals(sorted[k], selector.selectSPN(x, k, null), () -> "k[" + k + "]");
            Arrays.sort(x);
            Assertions.assertArrayEquals(sorted, x, () -> "Data destroyed: k[" + k + "]");
            if (k + 1 < sorted.length) {
                x = values.clone();
                Assertions.assertEquals(sorted[k], selector.selectSPN(x, k, kp1), () -> "k[" + k + "] with k+1");
                Assertions.assertEquals(sorted[k + 1], kp1[0], () -> "k+1[" + (k + 1) + "]");
                Arrays.sort(x);
                Assertions.assertArrayEquals(sorted, x, () -> "Data destroyed: k[" + k + "] with k+1");
            }
        }
    }
```

### Parameter provider — 同文件内的 `testSelect`

```java
    @ParameterizedTest
    @MethodSource
    void testSelect(double[] values) {
        final double[] sorted = values.clone();
        Arrays.sort(sorted);
        final KthSelector selector = new KthSelector();
        final double[] kp1 = new double[1];
        for (int i = 0; i < sorted.length; i++) {
            final int k = i;
            double[] x = values.clone();
            Assertions.assertEquals(sorted[k], selector.selectSP(x, k, null), () -> "k[" + k + "]");
            Arrays.sort(x);
            Assertions.assertArrayEquals(sorted, x, () -> "Data destroyed: k[" + k + "]");
            if (k + 1 < sorted.length) {
                x = values.clone();
                Assertions.assertEquals(sorted[k], selector.selectSP(x, k, kp1), () -> "k[" + k + "] with k+1");
                Assertions.assertEquals(sorted[k + 1], kp1[0], () -> "k+1[" + (k + 1) + "]");
                Arrays.sort(x);
                Assertions.assertArrayEquals(sorted, x, () -> "Data destroyed: k[" + k + "] with k+1");
            }
        }
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-024  ·  MethodSource

**项目** `SkyWalking`  **文件** `skywalking/oap-server/analyzer/meter-analyzer/src/test/java/org/apache/skywalking/oap/meter/analyzer/dsl/ScopeTest.java`  **测试** `test`

### 测试方法

```java
    @ParameterizedTest(name = "{0}")
    @MethodSource("data")
    public void test(final String name,
                     final ImmutableMap<String, SampleFamily> input,
                     final String expression,
                     final boolean isThrow,
                     final Map<MeterEntity, Sample[]> want) {
        Expression e = DSL.parse(name, expression);
        Result r = null;
        try {
            r = e.run(input);
        } catch (Throwable t) {
            if (isThrow) {
                return;
            }
            log.error("Test failed", t);
            fail("Should not throw anything");
        }
        if (isThrow) {
            fail("Should throw something");
        }
        assertThat(r.isSuccess()).isEqualTo(true);
        Map<MeterEntity, Sample[]> meterSamplesR = r.getData().context.getMeterSamples();
        meterSamplesR.forEach((meterEntity, samples) -> {
            assertThat(samples).isEqualTo(want.get(meterEntity));
        });
    }
```

### Parameter provider — 同文件内的 `data`

```java
    public static Collection<Object[]> data() {
        // This method is called before `@BeforeAll`.
        MeterEntity.setNamingControl(
            new NamingControl(512, 512, 512, new EndpointNameGrouping()));

        return Arrays.asList(new Object[][] {
            {
                "sum_service",
                of("http_success_request", SampleFamilyBuilder.newBuilder(
                    Sample.builder().labels(of("idc", "t1")).value(50).name("http_success_request").build(),
                    Sample.builder()
                          .labels(of("idc", "t3", "region", "cn", "svc", "catalog"))
                          .value(51)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t1", "region", "us", "svc", "product"))
                          .value(50)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t1", "region", "us", "instance", "10.0.0.1"))
                          .value(100)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t3", "region", "cn", "instance", "10.0.0.1"))
                          .value(3)
                          .name("http_success_request")
                          .build()
                ).build()),
                "http_success_request.sum(['idc']).service(['idc'], Layer.GENERAL)",
                false,
                new HashMap<MeterEntity, Sample[]>() {
                    {
                        put(
                            MeterEntity.newService("t1", Layer.GENERAL),
                            new Sample[] {Sample.builder().labels(of()).value(200).name("http_success_request").build()}
                        );
                        put(
                            MeterEntity.newService("t3", Layer.GENERAL),
                            new Sample[] {Sample.builder().labels(of()).value(54).name("http_success_request").build()}
                        );
                    }
                }
            },
            {
                "sum_service_labels",
                of("http_success_request", SampleFamilyBuilder.newBuilder(
                    Sample.builder().labels(of("idc", "t1")).value(50).name("http_success_request").build(),
                    Sample.builder()
                          .labels(of("idc", "t3", "region", "cn", "svc", "catalog"))
                          .value(51)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t1", "region", "us", "svc", "product"))
                          .value(50)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t1", "region", "us", "instance", "10.0.0.1"))
                          .value(100)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t3", "region", "cn", "instance", "10.0.0.1"))
                          .value(3)
                          .name("http_success_request")
                          .build()
                ).build()),
                "http_success_request.sum(['region', 'idc']).service(['idc'], Layer.GENERAL)",
                false,
                new HashMap<MeterEntity, Sample[]>() {
                    {
                        put(
                            MeterEntity.newService("t1", Layer.GENERAL),
                            new Sample[] {
                                Sample.builder()
                                      .labels(of("region", ""))
                                      .value(50)
                                      .name("http_success_request").build(),
                                Sample.builder()
                                      .labels(of("region", "us"))
                                      .value(150)
                                      .name("http_success_request").build()
                            }
                        );
                        put(
                            MeterEntity.newService("t3", Layer.GENERAL),
                            new Sample[] {
                                Sample.builder()
                                      .labels(of("region", "cn"))
                                      .value(54)
                                      .name("http_success_request").build()
                            }
                        );
                    }
                }
            },
            {
                "sum_service_m",
                of("http_success_request", SampleFamilyBuilder.newBuilder(
                    Sample.builder().labels(of("idc", "t1")).value(50).name("http_success_request").build(),
                    Sample.builder()
                          .labels(of("idc", "t3", "region", "cn", "svc", "catalog"))
                          .value(51)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t1", "region", "us", "svc", "product"))
                          .value(50)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t1", "region", "us", "instance", "10.0.0.1"))
                          .value(100)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t3", "region", "cn", "instance", "10.0.0.1"))
                          .value(3)
                          .name("http_success_request")
                          .build()
                ).build()),
                "http_success_request.sum(['idc', 'region']).service(['idc' , 'region'], Layer.GENERAL)",
                false,
                new HashMap<MeterEntity, Sample[]>() {
                    {
                        put(
                            MeterEntity.newService("t1.us", Layer.GENERAL),
                            new Sample[] {Sample.builder().labels(of()).value(150).name("http_success_request").build()}
                        );
                        put(
                            MeterEntity.newService("t3.cn", Layer.GENERAL),
                            new Sample[] {Sample.builder().labels(of()).value(54).name("http_success_request").build()}
                        );
                        put(
                            MeterEntity.newService("t1", Layer.GENERAL),
                            new Sample[] {Sample.builder().labels(of()).value(50).name("http_success_request").build()}
                        );
                    }
                }
            },
            {
                "sum_service_endpoint",
                of("http_success_request", SampleFamilyBuilder.newBuilder(
                    Sample.builder().labels(of("idc", "t1")).value(50).name("http_success_request").build(),
                    Sample.builder()
                          .labels(of("idc", "t3", "region", "cn", "svc", "catalog"))
                          .value(51)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t1", "region", "us", "svc", "product"))
                          .value(50)
                          .name("http_success_request")
                          .build(),
                    Sample.builder()
                          .labels(of("idc", "t1", "region", "us", "instance", "10.0.0.1"))
    // … 省略 394 行
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-027  ·  MethodSource

**项目** `PLC4X`  **文件** `plc4x/plc4j/drivers/opcua/src/test/java/org/apache/plc4x/java/opcua/OpcuaPlcDriverTest.java`  **测试** `readVariables`

### 测试方法

```java
        @ParameterizedTest
        @MethodSource("org.apache.plc4x.java.opcua.OpcuaPlcDriverTest#getConnectionSecurityPolicies")
        public void readVariables(SecurityPolicy policy, MessageSecurity messageSecurity) throws Exception {
            String connectionString = getConnectionString(policy, messageSecurity);
            PlcConnection opcuaConnection = new DefaultPlcDriverManager().getConnection(connectionString);
            Condition<PlcConnection> is_connected = new Condition<>(PlcConnection::isConnected, "is connected");
            assertThat(opcuaConnection).is(is_connected);

            PlcReadRequest.Builder builder = opcuaConnection.readRequestBuilder();
            tags.forEach((tagName, tagEntry) -> builder.addTagAddress(tagName, tagEntry.getKey()));
            PlcReadRequest request = builder.build();
            PlcReadResponse response = request.execute().get();

            SoftAssertions softly = new SoftAssertions();
            tags.keySet().forEach(tag -> {
                if (DOES_NOT_EXISTS_TAG_NAME.equals(tag)) {
                    softly.assertThat(response.getResponseCode(tag))
                        .describedAs("Tag %s should not exist and return NOT_FOUND status", tag)
                        .isEqualTo(PlcResponseCode.NOT_FOUND);
                } else {
                    softly.assertThat(response.getResponseCode(tag))
                        .describedAs("Tag %s should exist and return OK status", tag)
                        .isEqualTo(PlcResponseCode.OK);
                }
            });
            softly.assertAll();

            opcuaConnection.close();
            assertThat(opcuaConnection.isConnected()).isFalse();
        }
```

### Parameter provider — `OpcuaPlcDriverTest#getConnectionSecurityPolicies`（plc4x/plc4j/drivers/opcua/src/test/java/org/apache/plc4x/java/opcua/OpcuaPlcDriverTest.java）

```java

    private static Stream<Arguments> getConnectionSecurityPolicies() {
        return Stream.of(
            Arguments.of(SecurityPolicy.NONE, MessageSecurity.NONE),
            Arguments.of(SecurityPolicy.NONE, MessageSecurity.SIGN),
            Arguments.of(SecurityPolicy.NONE, MessageSecurity.SIGN_ENCRYPT),
            //Arguments.of(SecurityPolicy.Basic256Sha256, MessageSecurity.NONE),
            Arguments.of(SecurityPolicy.Basic256Sha256, MessageSecurity.SIGN),
            Arguments.of(SecurityPolicy.Basic256Sha256, MessageSecurity.SIGN_ENCRYPT),
            //Arguments.of(SecurityPolicy.Basic256, MessageSecurity.NONE),
            Arguments.of(SecurityPolicy.Basic256, MessageSecurity.SIGN),
            Arguments.of(SecurityPolicy.Basic256, MessageSecurity.SIGN_ENCRYPT),
            //Arguments.of(SecurityPolicy.Basic128Rsa15, MessageSecurity.NONE),
            Arguments.of(SecurityPolicy.Basic128Rsa15, MessageSecurity.SIGN),
            Arguments.of(SecurityPolicy.Basic128Rsa15, MessageSecurity.SIGN_ENCRYPT),
            //Arguments.of(SecurityPolicy.Aes128_Sha256_RsaOaep, MessageSecurity.NONE),
            Arguments.of(SecurityPolicy.Aes128_Sha256_RsaOaep, MessageSecurity.SIGN),
            Arguments.of(SecurityPolicy.Aes128_Sha256_RsaOaep, MessageSecurity.SIGN_ENCRYPT),
            //Arguments.of(SecurityPolicy.Aes256_Sha256_RsaPss, MessageSecurity.NONE),
            Arguments.of(SecurityPolicy.Aes256_Sha256_RsaPss, MessageSecurity.SIGN),
            Arguments.of(SecurityPolicy.Aes256_Sha256_RsaPss, MessageSecurity.SIGN_ENCRYPT)
        );
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-030  ·  ValueSource

**项目** `Commons-Numbers`  **文件** `commons-numbers/commons-numbers-gamma/src/test/java/org/apache/commons/numbers/gamma/BoostGammaTest.java`  **测试** `testGammaQLargeX`

### 测试方法

```java
    @ParameterizedTest
    @ValueSource(strings = {"igamma_int_data.csv", "igamma_med_data.csv", "igamma_big_data.csv"})
    @Order(1)
    void testGammaQLargeX(String datafile) throws Exception {
        assertIgammaLargeX("Commons", datafile, true, getLargeXTarget(), getUseAsymApprox(), true);
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-033  ·  EnumSource

**项目** `ZooKeeper`  **文件** `zookeeper/zookeeper-server/src/test/java/org/apache/zookeeper/server/admin/CommandAuthTest.java`  **测试** `testAuthCheck_authorized`

### 测试方法

```java
    @ParameterizedTest
    @EnumSource(AuthSchema.class)
    public void testAuthCheck_authorized(final AuthSchema authSchema) throws Exception {
        setupRootACL(authSchema);
        try {
            final HttpURLConnection authTestConn = sendAuthTestCommandRequest(authSchema, true);
            assertEquals(HttpURLConnection.HTTP_OK, authTestConn.getResponseCode());
        } finally {
            addAuthInfo(zk, authSchema);
            resetRootACL(zk);
        }
    }
```

### 枚举声明 — `AuthSchema`（zookeeper/zookeeper-server/src/test/java/org/apache/zookeeper/server/admin/CommandAuthTest.java）

```java
    public enum AuthSchema {
        DIGEST,
        X509,
        IP
    }
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-036  ·  EnumSource

**项目** `Avro`  **文件** `avro/lang/java/avro/src/test/java/org/apache/avro/TestReadingWritingDataInEvolvedSchemas.java`  **测试** `doubleWrittenWithUnionSchemaIsNotConvertedToFloatSchema`

### 测试方法

```java
  @ParameterizedTest
  @EnumSource(EncoderType.class)
  void doubleWrittenWithUnionSchemaIsNotConvertedToFloatSchema(EncoderType encoderType) throws Exception {
    Schema writer = UNION_INT_LONG_FLOAT_DOUBLE_RECORD;
    Record record = defaultRecordWithSchema(writer, FIELD_A, 42.0);
    byte[] encoded = encodeGenericBlob(record, encoderType);
    AvroTypeException exception = Assertions.assertThrows(AvroTypeException.class,
        () -> decodeGenericBlob(FLOAT_RECORD, writer, encoded, encoderType));
    Assertions.assertEquals("Found double, expecting float", exception.getMessage());
  }
```

### 枚举声明 — `EncoderType`（avro/lang/java/avro/src/test/java/org/apache/avro/TestReadingWritingDataInEvolvedSchemas.java）

```java
  enum EncoderType {
    BINARY, JSON
  }
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-039  ·  MethodSource

**项目** `Calcite`  **文件** `calcite/testkit/src/main/java/org/apache/calcite/test/QuidemTest.java`  **测试** `test`

### 测试方法

```java
  @ParameterizedTest
  @MethodSource("getPath")
  public void test(String path) throws Exception {
    final Method method = findMethod(path);
    if (method != null) {
      try {
        method.invoke(this, path);
      } catch (InvocationTargetException e) {
        Throwable cause = e.getCause();
        if (cause instanceof Exception) {
          throw (Exception) cause;
        }
        if (cause instanceof Error) {
          throw (Error) cause;
        }
        throw e;
      }
    } else {
      checkRun(path);
    }
  }
```

### Parameter provider — 同文件内的 `getPath`

```java
  protected abstract Collection<String> getPath();

  /** Quidem connection factory for Calcite's built-in test schemas. */
  protected static class QuidemConnectionFactory
      implements Quidem.ConnectionFactory {
    public Connection connect(String name) throws Exception {
      return connect(name, false);
    }

    @Override public Connection connect(String name, boolean reference)
        throws Exception {
      if (reference) {
        if (name.equals("foodmart")) {
          final ConnectionSpec db =
              CalciteAssert.DatabaseInstance.HSQLDB.foodmart;
          final Connection connection =
              DriverManager.getConnection(db.url, db.username,
                  db.password);
          connection.setSchema("foodmart");
          return connection;
        }
        return null;
      }
      switch (name) {
      case "hr":
        return CalciteAssert.hr()
            .connect();
      case "aux":
        return CalciteAssert.hr()
            .with(CalciteAssert.Config.AUX)
            .connect();
      case "foodmart":
        return CalciteAssert.that()
            .with(CalciteAssert.Config.FOODMART_CLONE)
            .connect();
      case "geo":
        return CalciteAssert.that()
            .with(CalciteAssert.Config.GEO)
            .connect();
      case "scott":
        return CalciteAssert.that()
            .with(CalciteAssert.Config.SCOTT)
            .connect();
      case "jdbc_scott":
        return CalciteAssert.that()
            .with(CalciteAssert.Config.JDBC_SCOTT)
            .connect();
      case "steelwheels":
        return CalciteAssert.that()
            .with(CalciteAssert.SchemaSpec.STEELWHEELS)
            .connect();
      case "jdbc_steelwheels":
        return CalciteAssert.that()
            .with(CalciteAssert.SchemaSpec.JDBC_STEELWHEELS)
            .connect();
      case "post":
        return CalciteAssert.that()
            .with(CalciteAssert.Config.REGULAR)
            .with(CalciteAssert.SchemaSpec.POST)
            .connect();
      case "post-postgresql":
        return CalciteAssert.that()
            .with(CalciteConnectionProperty.FUN, "standard,postgresql")
            .with(CalciteAssert.Config.REGULAR)
            .with(CalciteAssert.SchemaSpec.POST)
            .connect();
      case "post-big-query":
        return CalciteAssert.that()
            .with(CalciteConnectionProperty.FUN, "standard,bigquery")
            .with(CalciteAssert.Config.REGULAR)
            .with(CalciteAssert.SchemaSpec.POST)
            .connect();
      case "mysqlfunc":
        return CalciteAssert.that()
            .with(CalciteConnectionProperty.FUN, "mysql")
            .with(CalciteAssert.Config.REGULAR)
            .with(CalciteAssert.SchemaSpec.POST)
            .connect();
      case "sparkfunc":
        return CalciteAssert.that()
            .with(CalciteConnectionProperty.FUN, "spark")
            .with(CalciteAssert.Config.REGULAR)
            .with(CalciteAssert.SchemaSpec.POST)
            .connect();
      case "oraclefunc":
        return CalciteAssert.that()
            .with(CalciteConnectionProperty.FUN, "oracle")
            .with(CalciteAssert.Config.REGULAR)
            .connect();
      case "mssqlfunc":
        return CalciteAssert.that()
            .with(CalciteConnectionProperty.FUN, "mssql")
            .with(CalciteAssert.Config.REGULAR)
            .connect();
      case "catchall":
        return CalciteAssert.that()
            .with(CalciteConnectionProperty.TIME_ZONE, "UTC")
            .withSchema("s",
                new ReflectiveSchemaWithoutRowCount(
                    new CatchallSchema()))
            .connect();
      case "orinoco":
        return CalciteAssert.that()
            .with(CalciteAssert.SchemaSpec.ORINOCO)
            .connect();
      case "seq":
        final Connection connection = CalciteAssert.that()
            .withSchema("s", new AbstractSchema())
            .connect();
        connection.unwrap(CalciteConnection.class).getRootSchema()
            .subSchemas().get("s")
            .add("my_seq",
                new AbstractTable() {
                  @Override public RelDataType getRowType(
                      RelDataTypeFactory typeFactory) {
                    return typeFactory.builder()
                        .add("$seq", SqlTypeName.BIGINT).build();
                  }

                  @Override public Schema.TableType getJdbcTableType() {
                    return Schema.TableType.SEQUENCE;
                  }
                });
        return connection;
      case "bookstore":
        return CalciteAssert.that()
            .with(CalciteAssert.SchemaSpec.BOOKSTORE)
            .connect();
      default:
        throw new RuntimeException("unknown connection '" + name + "'");
      }
    }
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-042  ·  MethodSource

**项目** `Commons-Pool`  **文件** `commons-pool/src/test/java/org/apache/commons/pool3/impl/CallStackTest.java`  **测试** `testPrintFilledStackTrace`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource("data")
    void testPrintFilledStackTrace(final CallStack stack) {
        stack.fillInStackTrace();
        stack.printStackTrace(new PrintWriter(writer));
        final String stackTrace = writer.toString();
        assertTrue(stackTrace.contains(getClass().getName()));
    }
```

### Parameter provider — 同文件内的 `data`

```java

    public static Stream<Arguments> data() {
        // @formatter:off
        return Stream.of(
                Arguments.arguments(new ThrowableCallStack("Test", false)),
                Arguments.arguments(new ThrowableCallStack("yyyy-MM-dd'T'HH:mm:ss.SSSXXX", true)),
                Arguments.arguments(new SecurityManagerCallStack("Test", false)),
                Arguments.arguments(new SecurityManagerCallStack("yyyy-MM-dd'T'HH:mm:ss.SSSXXX", true))
        );
        // @formatter:on
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-045  ·  MethodSource

**项目** `Commons-CLI`  **文件** `commons-cli/src/test/java/org/apache/commons/cli/TypeHandlerTest.java`  **测试** `testCreateValue`

### 测试方法

```java
    @SuppressWarnings("unchecked")
    @ParameterizedTest(name = "{0} as {1}")
    @MethodSource("createValueTestParameters")
    void testCreateValue(final String str, final Class<?> type, final Object expected) throws Exception {
        @SuppressWarnings("cast")
        final Object objectApiTest = type; // KEEP this cast
        if (expected instanceof Class<?> && Throwable.class.isAssignableFrom((Class<?>) expected)) {
            assertThrows((Class<Throwable>) expected, () -> TypeHandler.createValue(str, type));
            assertThrows((Class<Throwable>) expected, () -> TypeHandler.createValue(str, objectApiTest));
        } else {
            assertEquals(expected, TypeHandler.createValue(str, type));
            assertEquals(expected, TypeHandler.createValue(str, objectApiTest));
        }
    }
```

### Parameter provider — 同文件内的 `createValueTestParameters`

```java

    private static Stream<Arguments> createValueTestParameters() throws MalformedURLException {
        // force the PatternOptionBuilder to load / modify the TypeHandler table.
        @SuppressWarnings("unused")
        final Class<?> loadStatic = PatternOptionBuilder.FILES_VALUE;
        // reset the type handler table.
        // TypeHandler.resetConverters();
        final List<Arguments> list = new ArrayList<>();

        /*
         * Dates calculated from strings are dependent upon configuration and environment settings for the machine on which the test is running. To avoid this
         * problem, convert the time into a string and then unparse that using the converter. This produces strings that always match the correct time zone.
         */
        final Date date = new Date(1023400137000L);
        final DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy");

        list.add(Arguments.of(Instantiable.class.getName(), PatternOptionBuilder.CLASS_VALUE, Instantiable.class));
        list.add(Arguments.of("what ever", PatternOptionBuilder.CLASS_VALUE, ParseException.class));

        list.add(Arguments.of("what ever", PatternOptionBuilder.DATE_VALUE, ParseException.class));
        list.add(Arguments.of(dateFormat.format(date), PatternOptionBuilder.DATE_VALUE, date));
        list.add(Arguments.of("Jun 06 17:48:57 EDT 2002", PatternOptionBuilder.DATE_VALUE, ParseException.class));

        list.add(Arguments.of("non-existing.file", PatternOptionBuilder.EXISTING_FILE_VALUE, ParseException.class));

        list.add(Arguments.of("some-file.txt", PatternOptionBuilder.FILE_VALUE, new File("some-file.txt")));

        list.add(Arguments.of("some-path.txt", Path.class, new File("some-path.txt").toPath()));

        // the PatternOptionBuilder.FILES_VALUE is not registered so it should just return the string
        list.add(Arguments.of("some.files", PatternOptionBuilder.FILES_VALUE, "some.files"));

        list.add(Arguments.of("just-a-string", Integer.class, ParseException.class));
        list.add(Arguments.of("5", Integer.class, 5));
        list.add(Arguments.of("5.5", Integer.class, ParseException.class));
        list.add(Arguments.of(Long.toString(Long.MAX_VALUE), Integer.class, ParseException.class));

        list.add(Arguments.of("just-a-string", Long.class, ParseException.class));
        list.add(Arguments.of("5", Long.class, 5L));
        list.add(Arguments.of("5.5", Long.class, ParseException.class));

        list.add(Arguments.of("just-a-string", Short.class, ParseException.class));
        list.add(Arguments.of("5", Short.class, (short) 5));
        list.add(Arguments.of("5.5", Short.class, ParseException.class));
        list.add(Arguments.of(Integer.toString(Integer.MAX_VALUE), Short.class, ParseException.class));

        list.add(Arguments.of("just-a-string", Byte.class, ParseException.class));
        list.add(Arguments.of("5", Byte.class, (byte) 5));
        list.add(Arguments.of("5.5", Byte.class, ParseException.class));
        list.add(Arguments.of(Short.toString(Short.MAX_VALUE), Byte.class, ParseException.class));

        list.add(Arguments.of("just-a-string", Character.class, 'j'));
        list.add(Arguments.of("5", Character.class, '5'));
        list.add(Arguments.of("5.5", Character.class, '5'));
        list.add(Arguments.of("\\u0124", Character.class, Character.toChars(0x0124)[0]));

        list.add(Arguments.of("just-a-string", Double.class, ParseException.class));
        list.add(Arguments.of("5", Double.class, 5d));
        list.add(Arguments.of("5.5", Double.class, 5.5));

        list.add(Arguments.of("just-a-string", Float.class, ParseException.class));
        list.add(Arguments.of("5", Float.class, 5f));
        list.add(Arguments.of("5.5", Float.class, 5.5f));
        list.add(Arguments.of(Double.toString(Double.MAX_VALUE), Float.class, Float.POSITIVE_INFINITY));

        list.add(Arguments.of("just-a-string", BigInteger.class, ParseException.class));
        list.add(Arguments.of("5", BigInteger.class, new BigInteger("5")));
        list.add(Arguments.of("5.5", BigInteger.class, ParseException.class));

        list.add(Arguments.of("just-a-string", BigDecimal.class, ParseException.class));
        list.add(Arguments.of("5", BigDecimal.class, new BigDecimal("5")));
        list.add(Arguments.of("5.5", BigDecimal.class, new BigDecimal(5.5)));

        list.add(Arguments.of("1.5", PatternOptionBuilder.NUMBER_VALUE, Double.valueOf(1.5)));
        list.add(Arguments.of("15", PatternOptionBuilder.NUMBER_VALUE, Long.valueOf(15)));
        list.add(Arguments.of("not a number", PatternOptionBuilder.NUMBER_VALUE, ParseException.class));

        list.add(Arguments.of(Instantiable.class.getName(), PatternOptionBuilder.OBJECT_VALUE, new Instantiable()));
        list.add(Arguments.of(NotInstantiable.class.getName(), PatternOptionBuilder.OBJECT_VALUE, ParseException.class));
        list.add(Arguments.of("unknown", PatternOptionBuilder.OBJECT_VALUE, ParseException.class));

        list.add(Arguments.of("String", PatternOptionBuilder.STRING_VALUE, "String"));

        final String urlString = "https://commons.apache.org";
        list.add(Arguments.of(urlString, PatternOptionBuilder.URL_VALUE, new URL(urlString)));
        list.add(Arguments.of("Malformed-url", PatternOptionBuilder.URL_VALUE, ParseException.class));

        return list.stream();

    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-048  ·  MethodSource

**项目** `Hive`  **文件** `hive/ql/src/test/org/apache/hadoop/hive/ql/io/parquet/serde/TestParquetTimestampsHive2Compatibility.java`  **测试** `testWriteHive2ReadHive4UsingLegacyConversionWithZone`

### 测试方法

```java
  @ParameterizedTest(name = "{0}")
  @MethodSource("generateTimestamps")
  void testWriteHive2ReadHive4UsingLegacyConversionWithZone(String timestampString) {
    TimeZone original = TimeZone.getDefault();
    try {
      String zoneId = "US/Pacific";
      TimeZone.setDefault(TimeZone.getTimeZone(zoneId));
      NanoTime nt = writeHive2(timestampString);
      Timestamp ts = readHive4(nt, zoneId, true);
      assertEquals(timestampString, ts.toString());
    } finally {
      TimeZone.setDefault(original);
    }
  }
```

### Parameter provider — 同文件内的 `generateTimestamps`

```java

  private static Stream<String> generateTimestamps() {
    return Stream.concat(Stream.generate(new Supplier<String>() {
      int i = 0;

      @Override
      public String get() {
        StringBuilder sb = new StringBuilder(29);
        int year = (i % 9999) + 1;
        sb.append(zeros(4 - digits(year)));
        sb.append(year);
        sb.append('-');
        int month = (i % 12) + 1;
        sb.append(zeros(2 - digits(month)));
        sb.append(month);
        sb.append('-');
        int day = (i % 28) + 1;
        sb.append(zeros(2 - digits(day)));
        sb.append(day);
        sb.append(' ');
        int hour = i % 24;
        sb.append(zeros(2 - digits(hour)));
        sb.append(hour);
        sb.append(':');
        int minute = i % 60;
        sb.append(zeros(2 - digits(minute)));
        sb.append(minute);
        sb.append(':');
        int second = i % 60;
        sb.append(zeros(2 - digits(second)));
        sb.append(second);
        sb.append('.');
        // Bitwise OR with one to avoid times with trailing zeros
        int nano = (i % 1000000000) | 1;
        sb.append(zeros(9 - digits(nano)));
        sb.append(nano);
        i++;
        return sb.toString();
      }
    })
    // Exclude dates falling in the default Gregorian change date since legacy code does not handle that interval
    // gracefully. It is expected that these do not work well when legacy APIs are in use. 
    .filter(s -> !s.startsWith("1582-10"))
    .limit(3000), Stream.of("9999-12-31 23:59:59.999"));
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-051  ·  MethodSource

**项目** `Zeppelin`  **文件** `zeppelin/zeppelin-plugins/notebookrepo/gcs/src/test/java/org/apache/zeppelin/notebook/repo/GCSNotebookRepoTest.java`  **测试** `testSave_create`

### 测试方法

```java
  @ParameterizedTest
  @MethodSource("buckets")
  void testSave_create(String bucketName, Optional<String> basePath, String uriPath) throws Exception {
    zConf.setProperty(ConfVars.ZEPPELIN_NOTEBOOK_GCS_STORAGE_DIR.getVarName(), uriPath);
    this.notebookRepo = new GCSNotebookRepo(zConf, noteParser, storage);
    notebookRepo.save(runningNote, AUTH_INFO);
    // Output is saved
    assertThat(storage.readAllBytes(makeBlobId(runningNote.getId(), runningNote.getPath(), bucketName, basePath)))
        .isEqualTo(runningNote.toJson().getBytes("UTF-8"));
  }
```

### Parameter provider — 同文件内的 `buckets`

```java

  private static Stream<Arguments> buckets() {
    return Stream.of(
      Arguments.of("bucketname", Optional.empty(), "gs://bucketname"),
      Arguments.of("bucketname-with-slash", Optional.empty(), "gs://bucketname-with-slash/"),
      Arguments.of("bucketname", Optional.of("path/to/dir"), "gs://bucketname/path/to/dir"),
      Arguments.of("bucketname", Optional.of("trailing/slash"), "gs://bucketname/trailing/slash/"));
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-054  ·  CsvSource

**项目** `Commons-Statistics`  **文件** `commons-statistics/commons-statistics-inference/src/test/java/org/apache/commons/statistics/inference/HypergeomTest.java`  **测试** `testDistribution`

### 测试方法

```java
    void testDistribution(int n, int k, int m) {
        final HypergeometricDistribution d1 = HypergeometricDistribution.of(n, k, m);
        final Hypergeom d2 = new Hypergeom(n, k, m);
        final int lo = d1.getSupportLowerBound();
        final int hi = d2.getSupportUpperBound();
        Assertions.assertEquals(lo, d2.getSupportLowerBound(), "lower bound");
        Assertions.assertEquals(hi, d2.getSupportUpperBound(), "upper bound");
        // Out of bounds will throw
        Assertions.assertThrows(IndexOutOfBoundsException.class, () -> d2.pmf(-1), "pmf(-1)");
        Assertions.assertThrows(IndexOutOfBoundsException.class, () -> d2.pmf(hi + 1), "pmf(hi + 1)");
        // Out of bounds for the cumulative functions is supported
        Assertions.assertEquals(d1.cumulativeProbability(lo - 1), d2.cdf(lo - 1), "cdf(lo - 1)");
        Assertions.assertEquals(d1.cumulativeProbability(hi + 1), d2.cdf(hi + 1), "cdf(hi + 1)");
        Assertions.assertEquals(d1.survivalProbability(lo - 1), d2.sf(lo - 1), "sf(lo - 1)");
        Assertions.assertEquals(d1.survivalProbability(hi + 1), d2.sf(hi + 1), "sf(hi + 1)");
        for (int x = lo; x <= hi; x++) {
            Assertions.assertEquals(d1.probability(x), d2.pmf(x), "pmf");
            Assertions.assertEquals(d1.cumulativeProbability(x), d2.cdf(x), "cdf");
            Assertions.assertEquals(d1.survivalProbability(x), d2.sf(x), "sf");
        }
        // Test the mode is the highest pmf
        double p = d2.pmf(lo);
        for (int x = lo + 1;; x++) {
            final double p2 = d2.pmf(x);
            if (p2 > p) {
                p = p2;
            } else {
                Assertions.assertEquals(d2.getLowerMode(), x - 1, "lower mode");
                break;
            }
        }
        p = d2.pmf(hi);
        for (int x = hi - 1;; x--) {
            final double p2 = d2.pmf(x);
            if (p2 > p) {
                p = p2;
            } else {
                Assertions.assertEquals(d2.getUpperMode(), x + 1, "upper mode");
                break;
            }
        }
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-057  ·  ValueSource

**项目** `Ozone`  **文件** `ozone/hadoop-ozone/integration-test/src/test/java/org/apache/hadoop/ozone/om/TestOMDbCheckpointServletInodeBasedXfer.java`  **测试** `testWriteDBToArchive`

### 测试方法

```java
  @ParameterizedTest
  @ValueSource(booleans = {true, false})
  public void testWriteDBToArchive(boolean expectOnlySstFiles) throws Exception {
    setupMocks();
    Path dbDir = folder.resolve("db_data");
    Files.createDirectories(dbDir);
    // Create dummy files: one SST, one non-SST
    Path sstFile = dbDir.resolve("test.sst");
    Files.write(sstFile, "sst content".getBytes(StandardCharsets.UTF_8)); // Write some content to make it non-empty

    Path nonSstFile = dbDir.resolve("test.log");
    Files.write(nonSstFile, "log content".getBytes(StandardCharsets.UTF_8));
    Set<String> sstFilesToExclude = new HashSet<>();
    AtomicLong maxTotalSstSize = new AtomicLong(1000000); // Sufficient size
    Map<String, String> hardLinkFileMap = new java.util.HashMap<>();
    Path tmpDir = folder.resolve("tmp");
    Files.createDirectories(tmpDir);
    TarArchiveOutputStream mockArchiveOutputStream = mock(TarArchiveOutputStream.class);
    List<String> fileNames = new ArrayList<>();
    try (MockedStatic<Archiver> archiverMock = mockStatic(Archiver.class)) {
      archiverMock.when(() -> Archiver.linkAndIncludeFile(any(), any(), any(), any())).thenAnswer(invocation -> {
        // Get the actual mockArchiveOutputStream passed from writeDBToArchive
        TarArchiveOutputStream aos = invocation.getArgument(2);
        File sourceFile = invocation.getArgument(0);
        String fileId = invocation.getArgument(1);
        fileNames.add(sourceFile.getName());
        aos.putArchiveEntry(new TarArchiveEntry(sourceFile, fileId));
        aos.write(new byte[100], 0, 100); // Simulate writing
        aos.closeArchiveEntry();
        return 100L;
      });
      boolean success = omDbCheckpointServletMock.writeDBToArchive(
          sstFilesToExclude, dbDir, maxTotalSstSize, mockArchiveOutputStream,
              tmpDir, hardLinkFileMap, expectOnlySstFiles);
      assertTrue(success);
      verify(mockArchiveOutputStream, times(fileNames.size())).putArchiveEntry(any());
      verify(mockArchiveOutputStream, times(fileNames.size())).closeArchiveEntry();
      verify(mockArchiveOutputStream, times(fileNames.size())).write(any(byte[].class), anyInt(),
          anyInt()); // verify write was called once

      boolean containsNonSstFile = false;
      for (String fileName : fileNames) {
        if (expectOnlySstFiles) {
          assertTrue(fileName.endsWith(".sst"), "File is not an SST File");
        } else {
          containsNonSstFile = true;
        }
      }

      if (!expectOnlySstFiles) {
        assertTrue(containsNonSstFile, "SST File is not expected");
      }
    }
  }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-060  ·  MethodSource

**项目** `Hive`  **文件** `hive/ql/src/test/org/apache/hadoop/hive/ql/io/parquet/serde/TestParquetTimestampsHive2Compatibility.java`  **测试** `testWriteHive2ReadHive4UsingLegacyConversionWithJulianLeapYears`

### 测试方法

```java
  @ParameterizedTest(name = "{0}")
  @MethodSource("generateTimestampsAndZoneIds")
  void testWriteHive2ReadHive4UsingLegacyConversionWithJulianLeapYears(String timestampString, String zoneId) {
    TimeZone original = TimeZone.getDefault();
    try {
      TimeZone.setDefault(TimeZone.getTimeZone(zoneId));
      NanoTime nt = writeHive2(timestampString);
      Timestamp ts = readHive4(nt, zoneId, true);
      assertEquals(timestampString, ts.toString());
    } finally {
      TimeZone.setDefault(original);
    }
  }
```

### Parameter provider — 同文件内的 `generateTimestampsAndZoneIds`

```java
  private static Stream<Arguments> generateTimestampsAndZoneIds() {
    return generateJulianLeapYearTimestamps().flatMap(
        timestampString -> Stream.of("Asia/Singapore", "Pacific/Kiritimati", "Etc/GMT+12", "Pacific/Niue")
            .map(zoneId -> Arguments.of(timestampString, zoneId)));
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-063  ·  CsvSource

**项目** `POI`  **文件** `poi/poi-scratchpad/src/test/java/org/apache/poi/hwpf/converter/TestWordToHtmlConverter.java`  **测试** `testFile`

### 测试方法

```java
    void testFile(String file, String contains) throws Exception {
        boolean emulatePictureStorage = !file.contains("equation");

        String result = getHtmlText(file, emulatePictureStorage);
        assertNotNull(result);
        // starting with JDK 9 such unimportant whitespaces may be trimmed
        result = result.replace("</a> <span", "</a><span");

        for (String match : contains.split("\\|")) {
            if (match.startsWith("!")) {
                assertNotContained(result, match.substring(1));
            } else {
                assertContains(result, match);
            }
        }
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-066  ·  MethodSource

**项目** `Calcite`  **文件** `calcite/testkit/src/main/java/org/apache/calcite/test/SqlOperatorTest.java`  **测试** `testCastDecimalToDoubleToInteger`

### 测试方法

```java
  @ParameterizedTest
  @MethodSource("safeParameters")
  void testCastDecimalToDoubleToInteger(CastType castType, SqlOperatorFixture f) {
    f.setFor(SqlStdOperatorTable.CAST, VmName.EXPAND);

    f.checkScalar("cast( cast(1.25 as double) as integer)", 1, "INTEGER NOT NULL");
    f.checkScalar("cast( cast(-1.25 as double) as integer)", -1, "INTEGER NOT NULL");
    f.checkScalar("cast( cast(1.75 as double) as integer)", 1, "INTEGER NOT NULL");
    f.checkScalar("cast( cast(-1.75 as double) as integer)", -1, "INTEGER NOT NULL");
    f.checkScalar("cast( cast(1.5 as double) as integer)", 1, "INTEGER NOT NULL");
    f.checkScalar("cast( cast(-1.5 as double) as integer)", -1, "INTEGER NOT NULL");
  }
```

### Parameter provider — 同文件内的 `safeParameters`

```java
  @SuppressWarnings("unused")
  private Stream<Arguments> safeParameters() {
    SqlOperatorFixture f = fixture();
    SqlOperatorFixture f2 =
        SqlOperatorFixtures.safeCastWrapper(f.withLibrary(SqlLibrary.BIG_QUERY), "SAFE_CAST");
    SqlOperatorFixture f3 =
        SqlOperatorFixtures.safeCastWrapper(f.withLibrary(SqlLibrary.MSSQL), "TRY_CAST");
    return Stream.of(
        () -> new Object[] {CastType.CAST, f},
        () -> new Object[] {CastType.SAFE_CAST, f2},
        () -> new Object[] {CastType.TRY_CAST, f3});
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-069  ·  ValueSource

**项目** `Maven`  **文件** `maven/impl/maven-core/src/test/java/org/apache/maven/graph/FilteredProjectDependencyGraphTest.java`  **测试** `downstreamProjectsShouldBeCached`

### 测试方法

```java
    @ParameterizedTest
    @ValueSource(booleans = {true, false})
    void downstreamProjectsShouldBeCached(boolean transitive) {
        FilteredProjectDependencyGraph graph =
                new FilteredProjectDependencyGraph(projectDependencyGraph, List.of(aProject));

        when(projectDependencyGraph.getDownstreamProjects(bProject, transitive)).thenReturn(List.of(cProject));

        graph.getDownstreamProjects(bProject, transitive);
        graph.getDownstreamProjects(bProject, transitive);

        verify(projectDependencyGraph).getDownstreamProjects(bProject, transitive);
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-072  ·  ValueSource

**项目** `JAMES`  **文件** `james-project/server/protocols/webadmin/webadmin-mailbox/src/test/java/org/apache/james/webadmin/routes/DomainQuotaRoutesNoVirtualHostingTest.java`  **测试** `allDeleteEndpointsShouldReturnNotAllowed`

### 测试方法

```java
    void allDeleteEndpointsShouldReturnNotAllowed(String endpoint) {
        given()
            .delete(endpoint)
        .then()
            .statusCode(HttpStatus.METHOD_NOT_ALLOWED_405);
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-075  ·  MethodSource

**项目** `Commons-Compress`  **文件** `commons-compress/src/test/java/org/apache/commons/compress/changes/ChangeSetRawTypesTest.java`  **测试** `testDeletePlusAddSame`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource("org.apache.commons.compress.changes.TestFixtures#getOutputArchiveNames")
    @SuppressWarnings({ "unchecked", "rawtypes" })
    void testDeletePlusAddSame(final String archiverName) throws Exception {
        final Path inputPath = createArchive(archiverName);
        final File testTxt = getFile("test.txt");
        final Path result = Files.createTempFile("test", "." + archiverName);
        try {
            try (InputStream inputStream = Files.newInputStream(inputPath);
                    ArchiveInputStream archiveInputStream = factory.createArchiveInputStream(archiverName, inputStream);
                    OutputStream newOutputStream = Files.newOutputStream(result);
                    ArchiveOutputStream archiveOutputStream = factory.createArchiveOutputStream(archiverName, newOutputStream);
                    InputStream csInputStream = Files.newInputStream(testTxt.toPath())) {
                setLongFileMode(archiveOutputStream);
                final ChangeSet changes = new ChangeSet();
                changes.delete("test/test3.xml");
                archiveListDelete("test/test3.xml");
                // Add a file
                final ArchiveEntry entry = archiveOutputStream.createArchiveEntry(testTxt, "test/test3.xml");
                changes.add(entry, csInputStream);
                archiveList.add("test/test3.xml");
                new ChangeSetPerformer(changes).perform(archiveInputStream, archiveOutputStream);
            }
            // Checks
            try (BufferedInputStream buf = new BufferedInputStream(Files.newInputStream(result));
                    ArchiveInputStream in = factory.createArchiveInputStream(buf)) {
                final File check = checkArchiveContent(in, archiveList, false);
                final File test3xml = new File(check, "result/test/test3.xml");
                assertEquals(testTxt.length(), test3xml.length());

                try (BufferedReader reader = new BufferedReader(Files.newBufferedReader(test3xml.toPath()))) {
                    String str;
                    while ((str = reader.readLine()) != null) {
                        // All lines look like this
                        "111111111111111111111111111000101011".equals(str);
                    }
                }
                forceDelete(check);
            }
        } finally {
            forceDelete(result);
        }
    }
```

### Parameter provider — `TestFixtures#getOutputArchiveNames`（commons-compress/src/test/java/org/apache/commons/compress/changes/TestFixtures.java）

```java

    static Set<String> getOutputArchiveNames() {
        final Set<String> outputStreamArchiveNames = ArchiveStreamFactory.DEFAULT.getOutputStreamArchiveNames();
        outputStreamArchiveNames.remove(ArchiveStreamFactory.AR); // TODO BUG?
        outputStreamArchiveNames.remove(ArchiveStreamFactory.SEVEN_Z); // TODO Does not support streaming.
        return outputStreamArchiveNames;
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-078  ·  MethodSource

**项目** `Druid`  **文件** `druid/sql/src/test/java/org/apache/druid/sql/calcite/CalciteJoinQueryTest.java`  **测试** `testInnerJoinOnTwoInlineDataSourcesWithOuterWhere_withLeftDirectAccess`

### 测试方法

```java
  @DecoupledTestConfig(quidemReason = QuidemTestCaseReason.JOIN_LEFT_DIRECT_ACCESS)
  @MethodSource("provideQueryContexts")
  @ParameterizedTest(name = "{0}")
  public void testInnerJoinOnTwoInlineDataSourcesWithOuterWhere_withLeftDirectAccess(Map<String, Object> queryContext)
  {
    queryContext = withLeftDirectAccessEnabled(queryContext);
    testQuery(
        "with abc as\n"
        + "(\n"
        + "  SELECT dim1, \"__time\", m1 from foo WHERE \"dim1\" = '10.1'\n"
        + ")\n"
        + "SELECT t1.dim1, t1.\"__time\" from abc as t1 INNER JOIN abc as t2 on t1.dim1 = t2.dim1 WHERE t1.dim1 = '10.1'\n",
        queryContext,
        ImmutableList.of(
            newScanQueryBuilder()
                .dataSource(
                    join(
                        new TableDataSource(CalciteTests.DATASOURCE1),
                        new QueryDataSource(
                            newScanQueryBuilder()
                                .dataSource(CalciteTests.DATASOURCE1)
                                .intervals(querySegmentSpec(Filtration.eternity()))
                                .filters(equality("dim1", "10.1", ColumnType.STRING))
                                .columns(ImmutableList.of("dim1"))
                                .columnTypes(ColumnType.STRING)
                                .resultFormat(ScanQuery.ResultFormat.RESULT_FORMAT_COMPACTED_LIST)
                                .context(queryContext)
                                .build()
                        ),
                        "j0.",
                        equalsCondition(
                            makeExpression("'10.1'"),
                            makeColumnExpression("j0.dim1")
                        ),
                        JoinType.INNER,
                        equality("dim1", "10.1", ColumnType.STRING)
                    )
                )
                .intervals(querySegmentSpec(Filtration.eternity()))
                .virtualColumns(expressionVirtualColumn("v0", "\'10.1\'", ColumnType.STRING))
                .columns("v0", "__time")
                .columnTypes(ColumnType.STRING, ColumnType.LONG)
                .context(queryContext)
                .build()
        ),
        ImmutableList.of(
            new Object[]{"10.1", 946771200000L}
        )
    );
  }
```

### Parameter provider — `provideQueryContexts`（druid/sql/src/test/java/org/apache/druid/sql/calcite/BaseCalciteQueryTest.java）

```java
  public static Object[] provideQueryContexts()
  {
    return new Object[] {
        // default behavior
        Named.of("default", QUERY_CONTEXT_DEFAULT),
        // all rewrites enabled
        Named.of("all_enabled", new ImmutableMap.Builder<String, Object>()
            .putAll(QUERY_CONTEXT_DEFAULT)
            .put(QueryContexts.JOIN_FILTER_REWRITE_VALUE_COLUMN_FILTERS_ENABLE_KEY, true)
            .put(QueryContexts.JOIN_FILTER_REWRITE_ENABLE_KEY, true)
            .put(QueryContexts.REWRITE_JOIN_TO_FILTER_ENABLE_KEY, true)
            .build()),
        // filter-on-value-column rewrites disabled, everything else enabled
        Named.of("filter-on-value-column_disabled", new ImmutableMap.Builder<String, Object>()
            .putAll(QUERY_CONTEXT_DEFAULT)
            .put(QueryContexts.JOIN_FILTER_REWRITE_VALUE_COLUMN_FILTERS_ENABLE_KEY, false)
            .put(QueryContexts.JOIN_FILTER_REWRITE_ENABLE_KEY, true)
            .put(QueryContexts.REWRITE_JOIN_TO_FILTER_ENABLE_KEY, true)
            .build()),
        // filter rewrites fully disabled, join-to-filter enabled
        Named.of("join-to-filter", new ImmutableMap.Builder<String, Object>()
            .putAll(QUERY_CONTEXT_DEFAULT)
            .put(QueryContexts.JOIN_FILTER_REWRITE_VALUE_COLUMN_FILTERS_ENABLE_KEY, false)
            .put(QueryContexts.JOIN_FILTER_REWRITE_ENABLE_KEY, false)
            .put(QueryContexts.REWRITE_JOIN_TO_FILTER_ENABLE_KEY, true)
            .build()),
        // filter rewrites disabled, but value column filters still set to true
        // (it should be ignored and this should
        // behave the same as the previous context)
        Named.of("filter-rewrites-disabled", new ImmutableMap.Builder<String, Object>()
            .putAll(QUERY_CONTEXT_DEFAULT)
            .put(QueryContexts.JOIN_FILTER_REWRITE_VALUE_COLUMN_FILTERS_ENABLE_KEY, true)
            .put(QueryContexts.JOIN_FILTER_REWRITE_ENABLE_KEY, false)
            .put(QueryContexts.REWRITE_JOIN_TO_FILTER_ENABLE_KEY, true)
            .build()),
        // filter rewrites fully enabled, join-to-filter disabled
        Named.of("filter-rewrites", new ImmutableMap.Builder<String, Object>()
            .putAll(QUERY_CONTEXT_DEFAULT)
            .put(QueryContexts.JOIN_FILTER_REWRITE_VALUE_COLUMN_FILTERS_ENABLE_KEY, true)
            .put(QueryContexts.JOIN_FILTER_REWRITE_ENABLE_KEY, true)
            .put(QueryContexts.REWRITE_JOIN_TO_FILTER_ENABLE_KEY, false)
            .build()),
        // all rewrites disabled
        Named.of("all_disabled", new ImmutableMap.Builder<String, Object>()
            .putAll(QUERY_CONTEXT_DEFAULT)
            .put(QueryContexts.JOIN_FILTER_REWRITE_VALUE_COLUMN_FILTERS_ENABLE_KEY, false)
            .put(QueryContexts.JOIN_FILTER_REWRITE_ENABLE_KEY, false)
            .put(QueryContexts.REWRITE_JOIN_TO_FILTER_ENABLE_KEY, false)
            .build()),
    };
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-081  ·  ValueSource

**项目** `SeaTunnel`  **文件** `seatunnel/seatunnel-connectors-v2/connector-fake/src/test/java/org/apache/seatunnel/connectors/seatunnel/fake/source/FakeDataGeneratorTest.java`  **测试** `testAutoIncrementId`

### 测试方法

```java
    @ParameterizedTest
    @ValueSource(strings = {"fake-auto-increment-id.conf", "fake-auto-increment-id.conf"})
    public void testAutoIncrementId(String conf) throws FileNotFoundException, URISyntaxException {
        ReadonlyConfig testConfig = getTestConfigFile(conf);
        int parallelism = testConfig.getOptional(EnvCommonOptions.PARALLELISM).orElse(1);
        FakeConfig fakeConfig = FakeConfig.buildWithConfig(testConfig);
        List<CompletableFuture<List<SeaTunnelRow>>> futures = new ArrayList<>();
        String jobId = UUID.randomUUID().toString();
        for (int i = 0; i < parallelism; i++) {
            CompletableFuture<List<SeaTunnelRow>> uCompletableFuture =
                    CompletableFuture.supplyAsync(
                            () -> {
                                FakeDataGenerator fakeDataGenerator =
                                        new FakeDataGenerator(fakeConfig, jobId);
                                return fakeDataGenerator.generateFakedRows(fakeConfig.getRowNum());
                            });
            futures.add(uCompletableFuture);
        }
        CompletableFuture.allOf(futures.toArray(new CompletableFuture[0]));
        List<SeaTunnelRow> seaTunnelRows =
                futures.stream()
                        .map(CompletableFuture::join)
                        .flatMap(List::stream)
                        .collect(Collectors.toList());
        List<Integer> ids =
                seaTunnelRows.stream()
                        .map(seaTunnelRow -> (int) seaTunnelRow.getField(0))
                        .distinct()
                        .sorted(Integer::compareTo)
                        .collect(Collectors.toList());
        Assertions.assertEquals(200, ids.size());
        ids.stream().min(Integer::compareTo).ifPresent(min -> Assertions.assertEquals(100, min));
        ids.stream().max(Integer::compareTo).ifPresent(max -> Assertions.assertEquals(299, max));
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-084  ·  MethodSource

**项目** `Commons-CSV`  **文件** `commons-csv/src/test/java/org/apache/commons/csv/CSVDuplicateHeaderTest.java`  **测试** `testCSVFormat`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource(value = {"duplicateHeaderAllowsMissingColumnsNamesData"})
    void testCSVFormat(final DuplicateHeaderMode duplicateHeaderMode,
                              final boolean allowMissingColumnNames,
                              final boolean ignoreHeaderCase,
                              final String[] headers,
                              final boolean valid) {
        final CSVFormat.Builder builder =
            CSVFormat.DEFAULT.builder()
                             .setDuplicateHeaderMode(duplicateHeaderMode)
                             .setAllowMissingColumnNames(allowMissingColumnNames)
                             .setIgnoreHeaderCase(ignoreHeaderCase)
                             .setHeader(headers);
        if (valid) {
            final CSVFormat format = builder.get();
            Assertions.assertEquals(duplicateHeaderMode, format.getDuplicateHeaderMode(), "DuplicateHeaderMode");
            Assertions.assertEquals(allowMissingColumnNames, format.getAllowMissingColumnNames(), "AllowMissingColumnNames");
            Assertions.assertArrayEquals(headers, format.getHeader(), "Header");
        } else {
            Assertions.assertThrows(IllegalArgumentException.class, builder::get);
        }
    }
```

### Parameter provider — 同文件内的 `duplicateHeaderAllowsMissingColumnsNamesData`

```java
    static Stream<Arguments> duplicateHeaderAllowsMissingColumnsNamesData() {
        return duplicateHeaderData()
            .filter(arg -> Boolean.TRUE.equals(arg.get()[1]) && Boolean.FALSE.equals(arg.get()[2]))
            .flatMap(arg -> {
                // Return test case with flags as all true/false combinations
                final Object[][] data = new Object[4][];
                final Boolean[] flags = {Boolean.TRUE, Boolean.FALSE};
                int i = 0;
                for (final Boolean a : flags) {
                    for (final Boolean b : flags) {
                        data[i] = arg.get().clone();
                        data[i][1] = a;
                        data[i][2] = b;
                        i++;
                    }
                }
                return Arrays.stream(data).map(Arguments::of);
            });
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-087  ·  EnumSource

**项目** `Avro`  **文件** `avro/lang/java/avro/src/test/java/org/apache/avro/TestReadingWritingDataInEvolvedSchemas.java`  **测试** `longWrittenWithUnionSchemaIsConvertedToLongFloatUnionSchema`

### 测试方法

```java
  @ParameterizedTest
  @EnumSource(EncoderType.class)
  void longWrittenWithUnionSchemaIsConvertedToLongFloatUnionSchema(EncoderType encoderType) throws Exception {
    Schema writer = UNION_LONG_RECORD;
    Record record = defaultRecordWithSchema(writer, FIELD_A, 42L);
    byte[] encoded = encodeGenericBlob(record, encoderType);
    Record decoded = decodeGenericBlob(UNION_LONG_FLOAT_RECORD, writer, encoded, encoderType);
    assertEquals(42L, decoded.get(FIELD_A));
  }
```

### 枚举声明 — `EncoderType`（avro/lang/java/avro/src/test/java/org/apache/avro/TestReadingWritingDataInEvolvedSchemas.java）

```java
  enum EncoderType {
    BINARY, JSON
  }
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-090  ·  MethodSource

**项目** `Directory-Studio`  **文件** `directory-studio/tests/test.integration.ui/src/main/java/org/apache/directory/studio/test/integration/ui/ValueEditorTest.java`  **测试** `testGetStringOrBinaryValue`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource("data")
    public void testGetStringOrBinaryValue( String name, Data data ) throws Exception
    {
        setup( name, data );
        if ( data.expectedStringOrBinaryValue instanceof byte[] )
        {
            assertArrayEquals( ( byte[] ) data.expectedStringOrBinaryValue,
                ( byte[] ) editor.getStringOrBinaryValue( editor.getRawValue( value ) ) );
        }
        else
        {
            assertEquals( data.expectedStringOrBinaryValue,
                editor.getStringOrBinaryValue( editor.getRawValue( value ) ) );
        }
    }
```

### Parameter provider — 同文件内的 `data`

```java

    public static Stream<Arguments> data()
    {
        return Stream.of( new Object[][]
            {
                /*
                 * InPlaceTextValueEditor can handle string values and binary values that can be decoded as UTF-8.
                 */

                {
                    "InPlaceTextValueEditor - empty value",
                    Data.data().valueEditorClass( InPlaceTextValueEditor.class ).attribute( CN )
                        .rawValue( IValue.EMPTY_STRING_VALUE ).expectedRawValue( EMPTY_STRING )
                        .expectedDisplayValue( EMPTY_STRING ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( EMPTY_STRING ) },

                {
                    "InPlaceTextValueEditor - empty string",
                    Data.data().valueEditorClass( InPlaceTextValueEditor.class ).attribute( CN )
                        .rawValue( EMPTY_STRING ).expectedRawValue( EMPTY_STRING ).expectedDisplayValue( EMPTY_STRING )
                        .expectedHasValue( true ).expectedStringOrBinaryValue( EMPTY_STRING ) },

                {
                    "InPlaceTextValueEditor - ascii",
                    Data.data().valueEditorClass( InPlaceTextValueEditor.class ).attribute( CN ).rawValue( ASCII )
                        .expectedRawValue( ASCII ).expectedDisplayValue( ASCII ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( ASCII ) },

                {
                    "InPlaceTextValueEditor - unicode",
                    Data.data().valueEditorClass( InPlaceTextValueEditor.class ).attribute( CN ).rawValue( UNICODE )
                        .expectedRawValue( UNICODE ).expectedDisplayValue( UNICODE ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( UNICODE ) },

                {
                    "InPlaceTextValueEditor - bytearray UTF8",
                    Data.data().valueEditorClass( InPlaceTextValueEditor.class ).attribute( USER_PWD ).rawValue( UTF8 )
                        .expectedRawValue( UNICODE ).expectedDisplayValue( UNICODE ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( UNICODE ) },

                // text editor always tries to decode byte[] as UTF-8, so it can not handle ISO-8859-1 encoded byte[] 
                {
                    "InPlaceTextValueEditor - bytearray ISO-8859-1",
                    Data.data().valueEditorClass( InPlaceTextValueEditor.class ).attribute( USER_PWD )
                        .rawValue( ISO88591 ).expectedRawValue( null ).expectedDisplayValue( IValueEditor.NULL )
                        .expectedHasValue( true ).expectedStringOrBinaryValue( null ) },

                // text editor always tries to decode byte[] as UTF-8, so it can not handle arbitrary byte[] 
                {
                    "InPlaceTextValueEditor - bytearray PNG",
                    Data.data().valueEditorClass( InPlaceTextValueEditor.class ).attribute( USER_PWD ).rawValue( PNG )
                        .expectedRawValue( null ).expectedDisplayValue( IValueEditor.NULL ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( null ) },

                /*
                 * InPlaceBooleanValueEditor can only handle TRUE or FALSE values.
                 */

                {
                    "InPlaceBooleanValueEditor - TRUE",
                    Data.data().valueEditorClass( InPlaceBooleanValueEditor.class ).attribute( CN ).rawValue( TRUE )
                        .expectedRawValue( TRUE ).expectedDisplayValue( TRUE ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( TRUE ) },

                {
                    "InPlaceBooleanValueEditor - FALSE",
                    Data.data().valueEditorClass( InPlaceBooleanValueEditor.class ).attribute( CN ).rawValue( FALSE )
                        .expectedRawValue( FALSE ).expectedDisplayValue( FALSE ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( FALSE ) },

                {
                    "InPlaceBooleanValueEditor - INVALID",
                    Data.data().valueEditorClass( InPlaceBooleanValueEditor.class ).attribute( CN )
                        .rawValue( "invalid" ).expectedRawValue( null ).expectedDisplayValue( IValueEditor.NULL )
                        .expectedHasValue( true ).expectedStringOrBinaryValue( null ) },

                {
                    "InPlaceBooleanValueEditor - bytearray TRUE",
                    Data.data().valueEditorClass( InPlaceBooleanValueEditor.class ).attribute( USER_PWD )
                        .rawValue( TRUE.getBytes( UTF_8 ) ).expectedRawValue( TRUE ).expectedDisplayValue( TRUE )
                        .expectedHasValue( true ).expectedStringOrBinaryValue( TRUE ) },

                {
                    "InPlaceBooleanValueEditor - bytearray FALSE",
                    Data.data().valueEditorClass( InPlaceBooleanValueEditor.class ).attribute( USER_PWD )
                        .rawValue( FALSE.getBytes( UTF_8 ) ).expectedRawValue( FALSE ).expectedDisplayValue( FALSE )
                        .expectedHasValue( true ).expectedStringOrBinaryValue( FALSE ) },

                {
                    "InPlaceBooleanValueEditor - bytearray INVALID",
                    Data.data().valueEditorClass( InPlaceBooleanValueEditor.class ).attribute( USER_PWD )
                        .rawValue( "invalid".getBytes( UTF_8 ) ).expectedRawValue( null )
                        .expectedDisplayValue( IValueEditor.NULL ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( null ) },

                /*
                 * InPlaceOidValueEditor can only handle OIDs
                 */

                {
                    "InPlaceOidValueEditor - numeric OID",
                    Data.data().valueEditorClass( InPlaceOidValueEditor.class ).attribute( CN ).rawValue( NUMERIC_OID )
                        .expectedRawValue( NUMERIC_OID ).expectedDisplayValue( NUMERIC_OID + " (Start TLS)" )
                        .expectedHasValue( true ).expectedStringOrBinaryValue( NUMERIC_OID ) },

                {
                    "InPlaceOidValueEditor - descr OID",
                    Data.data().valueEditorClass( InPlaceOidValueEditor.class ).attribute( CN ).rawValue( DESCR_OID )
                        .expectedRawValue( DESCR_OID ).expectedDisplayValue( DESCR_OID ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( DESCR_OID ) },

                {
                    "InPlaceOidValueEditor - relaxed descr OID",
                    Data.data().valueEditorClass( InPlaceOidValueEditor.class ).attribute( CN )
                        .rawValue( "orclDBEnterpriseRole_82" ).expectedRawValue( "orclDBEnterpriseRole_82" )
                        .expectedDisplayValue( "orclDBEnterpriseRole_82" ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( "orclDBEnterpriseRole_82" ) },

                {
                    "InPlaceOidValueEditor - INVALID",
                    Data.data().valueEditorClass( InPlaceOidValueEditor.class ).attribute( CN ).rawValue( "in valid" )
                        .expectedRawValue( null ).expectedDisplayValue( IValueEditor.NULL ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( null ) },

                {
                    "InPlaceOidValueEditor - bytearray numeric OID",
                    Data.data().valueEditorClass( InPlaceOidValueEditor.class ).attribute( USER_PWD )
                        .rawValue( NUMERIC_OID.getBytes( UTF_8 ) ).expectedRawValue( NUMERIC_OID )
                        .expectedDisplayValue( NUMERIC_OID + " (Start TLS)" ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( NUMERIC_OID ) },

                {
                    "InPlaceOidValueEditor - bytearray INVALID",
                    Data.data().valueEditorClass( InPlaceOidValueEditor.class ).attribute( USER_PWD )
                        .rawValue( "in valid".getBytes( UTF_8 ) ).expectedRawValue( null )
                        .expectedDisplayValue( IValueEditor.NULL ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( null ) },

                /*
                 * TextValueEditor can handle string values and binary values that can be decoded as UTF-8.
                 */

                {
                    "TextValueEditor - empty string value",
                    Data.data().valueEditorClass( TextValueEditor.class ).attribute( CN )
                        .rawValue( IValue.EMPTY_STRING_VALUE ).expectedRawValue( EMPTY_STRING )
                        .expectedDisplayValue( EMPTY_STRING ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( EMPTY_STRING ) },

                {
                    "TextValueEditor - empty string",
                    Data.data().valueEditorClass( TextValueEditor.class ).attribute( CN ).rawValue( EMPTY_STRING )
                        .expectedRawValue( EMPTY_STRING ).expectedDisplayValue( EMPTY_STRING ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( EMPTY_STRING ) },

                {
                    "TextValueEditor - ascii",
                    Data.data().valueEditorClass( TextValueEditor.class ).attribute( CN ).rawValue( ASCII )
                        .expectedRawValue( ASCII ).expectedDisplayValue( ASCII ).expectedHasValue( true )
                        .expectedStringOrBinaryValue( ASCII ) },
    // … 省略 106 行
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-093  ·  ValueSource

**项目** `Flink`  **文件** `flink/flink-table/flink-table-planner/src/test/java/org/apache/flink/table/planner/operations/SqlOtherOperationConverterTest.java`  **测试** `testHelpCommands`

### 测试方法

```java
    @ParameterizedTest
    @ValueSource(strings = {"HELP", "HELP;", "HELP ;", "HELP\t;", "HELP\n;"})
    void testHelpCommands(String command) {
        ExtendedParser extendedParser = new ExtendedParser();
        assertThat(extendedParser.parse(command)).get().isInstanceOf(HelpOperation.class);
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-096  ·  ValueSource

**项目** `Rat`  **文件** `creadur-rat/apache-rat-core/src/test/java/org/apache/rat/commandline/ArgTests.java`  **测试** `outputFleNameNoDirectoryTest`

### 测试方法

```java
    @ParameterizedTest(name = "{0}")
    @ValueSource(strings = { "rat.txt", "./rat.txt", "/rat.txt", "target/rat.test" })
    public void outputFleNameNoDirectoryTest(String name) throws ParseException, IOException {
        class OutputFileConfig extends ReportConfiguration  {
            private File actual = null;
            @Override
            public void setOut(File file) {
                actual = file;
            }
        }
        String fileName = name.replace("/", DocumentName.FSInfo.getDefault().dirSeparator());
        File expected = new File(fileName);

        CommandLine commandLine = createCommandLine(new String[] {"--output-file", fileName});
        OutputFileConfig configuration = new OutputFileConfig();
        ArgumentContext ctxt = new ArgumentContext(new File("."), configuration, commandLine);
        Arg.processArgs(ctxt);
        assertThat(configuration.actual.getAbsolutePath()).isEqualTo(expected.getCanonicalPath());
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-099  ·  MethodSource

**项目** `hadoop`  **文件** `hadoop/hadoop-cloud-storage-project/hadoop-tos/src/test/java/org/apache/hadoop/fs/tosfs/object/TestObjectStorage.java`  **测试** `testDeleteNonEmptyDir`

### 测试方法

```java
  @ParameterizedTest
  @MethodSource("provideArguments")
  public void testDeleteNonEmptyDir(ObjectStorage store) throws IOException {
    setEnv(store);
    storage.put("a/", new byte[0]);
    storage.put("a/b/", new byte[0]);
    assertArrayEquals(new byte[0], IOUtils.toByteArray(getStream("a/b/", 0, 256)));

    ListObjectsResponse response = list("a/b/", "a/b/", 100, "/");
    assertEquals(0, response.objects().size());
    assertEquals(0, response.commonPrefixes().size());

    if (!storage.bucket().isDirectory()) {
      // Directory bucket only supports list with delimiter = '/'.
      response = list("a/b/", "a/b/", 100, null);
      assertEquals(0, response.objects().size());
      assertEquals(0, response.commonPrefixes().size());
    }

    storage.delete("a/b/");
    assertNull(storage.head("a/b/"));
    assertNull(storage.head("a/b"));
    assertNotNull(storage.head("a/"));
  }
```

### Parameter provider — 同文件内的 `provideArguments`

```java

  public static Stream<Arguments> provideArguments() {
    assumeTrue(TestEnv.checkTestEnabled());

    List<Arguments> values = new ArrayList<>();
    for (ObjectStorage store : TestUtility.createTestObjectStorage(FILE_STORE_ROOT)) {
      values.add(Arguments.of(store));
    }
    return values.stream();
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-102  ·  MethodSource

**项目** `Commons-Statistics`  **文件** `commons-statistics/commons-statistics-descriptive/src/test/java/org/apache/commons/statistics/descriptive/BaseLongStatisticTest.java`  **测试** `testCombineEmpty`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource(value = "testAccept")
    final void testCombineEmpty(long[] values) {
        final S empty = create();
        final S nonEmpty = Statistics.add(create(), values);
        final StatisticResult expected = Statistics.add(create(), values);
        final S result = assertCombine(nonEmpty, empty);
        TestHelper.assertEquals(expected, result, null,
            () -> statisticName + " nonEmpty.combine(empty)");
    }
```

### Parameter provider — 同文件内的 `testAccept`

```java
    final Stream<Arguments> testAccept() {
        return streamSingleArrayData(getToleranceAccept(), StatisticTestData::getTolAccept);
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-105  ·  MethodSource

**项目** `Commons-JEXL`  **文件** `commons-jexl/src/test/java/org/apache/commons/jexl3/JXLTTest.java`  **测试** `test311i`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource("engines")
    void test311i(final JexlBuilder builder) {
        init(builder);
        final JexlContext ctx311 = new Context311();
        // @formatter:off
        final String rpt
                = "$$var u = 'Universe'; exec('4').execute((a, b)->{"
                + "\n<p>${u} ${a}${b}</p>"
                + "\n$$}, '2')";
        // @formatter:on
        final JxltEngine.Template t = JXLT.createTemplate("$$", new StringReader(rpt));
        final StringWriter strw = new StringWriter();
        t.evaluate(ctx311, strw, 42);
        final String output = strw.toString();
        assertEquals("<p>Universe 42</p>\n", output);
    }
```

### Parameter provider — 同文件内的 `engines`

```java

   public static List<JexlBuilder> engines() {
       final JexlFeatures f = new JexlFeatures();
       f.lexical(true).lexicalShade(true);
      return Arrays.asList(
              new JexlBuilder().silent(false).lexical(true).lexicalShade(true).cache(128).strict(true),
              new JexlBuilder().features(f).silent(false).cache(128).strict(true),
              new JexlBuilder().silent(false).cache(128).strict(true));
   }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-108  ·  EnumSource

**项目** `ZooKeeper`  **文件** `zookeeper/zookeeper-server/src/test/java/org/apache/zookeeper/server/quorum/LearnerSyncThrottlerTest.java`  **测试** `testParallelNoThrottle`

### 测试方法

```java
    @ParameterizedTest
    @EnumSource(LearnerSyncThrottler.SyncType.class)
    public void testParallelNoThrottle(LearnerSyncThrottler.SyncType syncType) {
        final int numThreads = 50;

        final LearnerSyncThrottler throttler = new LearnerSyncThrottler(numThreads, syncType);
        ExecutorService threadPool = Executors.newFixedThreadPool(numThreads);
        final CountDownLatch threadStartLatch = new CountDownLatch(numThreads);
        final CountDownLatch syncProgressLatch = new CountDownLatch(numThreads);

        List<Future<Boolean>> results = new ArrayList<>(numThreads);
        for (int i = 0; i < numThreads; i++) {
            results.add(threadPool.submit(new Callable<Boolean>() {

                @Override
                public Boolean call() {
                    threadStartLatch.countDown();
                    try {
                        threadStartLatch.await();

                        throttler.beginSync(false);

                        syncProgressLatch.countDown();
                        syncProgressLatch.await();

                        throttler.endSync();
                    } catch (Exception e) {
                        return false;
                    }

                    return true;
                }
            }));
        }

        try {
            for (Future<Boolean> result : results) {
                assertTrue(result.get());
            }
        } catch (Exception e) {

        } finally {
            threadPool.shutdown();
        }
    }
```

### 枚举声明 — `SyncType`（zookeeper/zookeeper-server/src/main/java/org/apache/zookeeper/server/quorum/LearnerSyncThrottler.java）

```java
    public enum SyncType {
        DIFF,
        SNAP
    }
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-111  ·  CsvSource

**项目** `Fineract`  **文件** `fineract/fineract-core/src/test/java/org/apache/fineract/util/LoopGuardTest.java`  **测试** `testSafeWhileLoopExecutesCorrectly`

### 测试方法

```java
    void testSafeWhileLoopExecutesCorrectly(int targetValue, int maxIterations) {
        TestContext context = new TestContext();

        Predicate<TestContext> condition = ctx -> ctx.iteration < targetValue;
        LoopGuard.LoopBody<TestContext> body = ctx -> ctx.iteration++;

        LoopGuard.runSafeWhileLoop(maxIterations, context, condition, body);

        Assertions.assertEquals(targetValue, context.iteration);
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-114  ·  ValueSource

**项目** `ORC`  **文件** `orc/java/mapreduce/src/test/org/apache/orc/mapred/TestOrcFileEvolution.java`  **测试** `testPreHive4243AddColumnWithFix`

### 测试方法

```java
  @ParameterizedTest
  @ValueSource(booleans = {true, false})
  public void testPreHive4243AddColumnWithFix(boolean addSarg) {
    checkEvolution("struct<_col0:int,_col1:string>",
                   "struct<a:int,b:string,c:double>",
                   struct(1, "foo"),
                   struct(1, "foo", null), true, addSarg, false);
  }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-117  ·  EnumSource

**项目** `Causeway`  **文件** `causeway/core/mmtest/src/test/java/org/apache/causeway/core/metamodel/valuesemantics/temporal/TemporalValueSemanticsProviderTest.java`  **测试** `timeFormats`

### 测试方法

```java
    @ParameterizedTest
    @EnumSource(TimePrecision.class)
    void timeFormats(final TimePrecision timePrecision) {

        target = new TemporalValueSemanticsProvider_forTesting(
                TemporalCharacteristic.TIME_ONLY, OffsetCharacteristic.LOCAL);

        Context context = null;
        LocalTime localTime = LocalTime.of(13, 12, 45);

        var formatter = target.getTemporalEditingFormat(context ,
                target.getTemporalCharacteristic(),
                target.getOffsetCharacteristic(),
                timePrecision,
                EditingFormatDirection.OUTPUT,
                editingPattern);

        var formattedTemporal = formatter.format(localTime);
        assertNotNull(formattedTemporal);
    }
```

### 枚举声明 — `TimePrecision`（causeway/api/applib/src/main/java/org/apache/causeway/applib/annotation/TimePrecision.java）

```java
public enum TimePrecision {

    UNSPECIFIED,

    /**
     * 9 fractional digits for <i>Second</i>
     */
    NANO_SECOND,

    /**
     * 6 fractional digits for <i>Second</i>
     */
    MICRO_SECOND,

    /**
     * 3 fractional digits for <i>Second</i>
     */
    MILLI_SECOND,

    /**
     * <i>Second</i>
     */
    SECOND,

    /**
     * <i>Minute</i>
     */
    MINUTE,

    /**
     * <i>Hour</i>
     */
    HOUR;

}
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-120  ·  EnumSource

**项目** `Commons-RNG`  **文件** `commons-rng/commons-rng-simple/src/test/java/org/apache/commons/rng/simple/internal/NativeSeedTypeParametricTest.java`  **测试** `testCreateSeed`

### 测试方法

```java
    @ParameterizedTest
    @EnumSource
    void testCreateSeed(NativeSeedType nativeSeedType) {
        final int size = 3;
        final Object seed = nativeSeedType.createSeed(size);
        Assertions.assertNotNull(seed);
        final Class<?> type = nativeSeedType.getType();
        Assertions.assertEquals(type, seed.getClass(), "Seed was not the correct class");
        if (type.isArray()) {
            Assertions.assertEquals(size, Array.getLength(seed), "Seed was not created the correct length");
        }
    }
```

### 枚举声明 — `NativeSeedType`（commons-rng/commons-rng-simple/src/main/java/org/apache/commons/rng/simple/internal/NativeSeedType.java）

```java
public enum NativeSeedType {
    /** The seed type is {@code Integer}. */
    INT(Integer.class, 4) {
        @Override
        public Integer createSeed(int size, int from, int to) {
            return SeedFactory.createInt();
        }
        @Override
        protected Integer convert(Integer seed, int size) {
            return seed;
        }
        @Override
        protected Integer convert(Long seed, int size) {
            return Conversions.long2Int(seed);
        }
        @Override
        protected Integer convert(int[] seed, int size) {
            return Conversions.intArray2Int(seed);
        }
        @Override
        protected Integer convert(long[] seed, int size) {
            return Conversions.longArray2Int(seed);
        }
        @Override
        protected Integer convert(byte[] seed, int size) {
            return Conversions.byteArray2Int(seed);
        }
    },
    /** The seed type is {@code Long}. */
    LONG(Long.class, 8) {
        @Override
        public Long createSeed(int size, int from, int to) {
            return SeedFactory.createLong();
        }
        @Override
        protected Long convert(Integer seed, int size) {
            return Conversions.int2Long(seed);
        }
        @Override
        protected Long convert(Long seed, int size) {
            return seed;
        }
        @Override
        protected Long convert(int[] seed, int size) {
            return Conversions.intArray2Long(seed);
        }
        @Override
        protected Long convert(long[] seed, int size) {
            return Conversions.longArray2Long(seed);
        }
        @Override
        protected Long convert(byte[] seed, int size) {
            return Conversions.byteArray2Long(seed);
        }
    },
    /** The seed type is {@code int[]}. */
    INT_ARRAY(int[].class, 4) {
        @Override
        public int[] createSeed(int size, int from, int to) {
            // Limit the number of calls to the synchronized method. The generator
            // will support self-seeding.
            return SeedFactory.createIntArray(Math.min(size, RANDOM_SEED_ARRAY_SIZE),
                                              from, to);
        }
        @Override
        protected int[] convert(Integer seed, int size) {
            return Conversions.int2IntArray(seed, size);
        }
        @Override
        protected int[] convert(Long seed, int size) {
            return Conversions.long2IntArray(seed, size);
        }
        @Override
        protected int[] convert(int[] seed, int size) {
            return seed;
        }
        @Override
        protected int[] convert(long[] seed, int size) {
            // Avoid zero filling seeds that are too short
            return Conversions.longArray2IntArray(seed,
                Math.min(size, Conversions.intSizeFromLongSize(seed.length)));
        }
        @Override
        protected int[] convert(byte[] seed, int size) {
            // Avoid zero filling seeds that are too short
            return Conversions.byteArray2IntArray(seed,
                Math.min(size, Conversions.intSizeFromByteSize(seed.length)));
        }
    },
    /** The seed type is {@code long[]}. */
    LONG_ARRAY(long[].class, 8) {
        @Override
        public long[] createSeed(int size, int from, int to) {
            // Limit the number of calls to the synchronized method. The generator
            // will support self-seeding.
            return SeedFactory.createLongArray(Math.min(size, RANDOM_SEED_ARRAY_SIZE),
                                               from, to);
        }
        @Override
        protected long[] convert(Integer seed, int size) {
            return Conversions.int2LongArray(seed, size);
        }
        @Override
        protected long[] convert(Long seed, int size) {
            return Conversions.long2LongArray(seed, size);
        }
        @Override
        protected long[] convert(int[] seed, int size) {
            // Avoid zero filling seeds that are too short
            return Conversions.intArray2LongArray(seed,
                Math.min(size, Conversions.longSizeFromIntSize(seed.length)));
        }
        @Override
        protected long[] convert(long[] seed, int size) {
            return seed;
        }
        @Override
        protected long[] convert(byte[] seed, int size) {
            // Avoid zero filling seeds that are too short
            return Conversions.byteArray2LongArray(seed,
                Math.min(size, Conversions.longSizeFromByteSize(seed.length)));
        }
    };

    /** Error message for unrecognized seed types. */
    private static final String UNRECOGNISED_SEED = "Unrecognized seed type: ";
    /** Maximum length of the seed array (for creating array seeds). */
    private static final int RANDOM_SEED_ARRAY_SIZE = 128;

    /** Define the class type of the native seed. */
    private final Class<?> type;

    /**
     * Define the number of bytes required to represent the native seed. If the type is
     * an array then this represents the size of a single value of the type.
     */
    private final int bytes;

    /**
     * Instantiates a new native seed type.
     *
     * @param type Define the class type of the native seed.
     * @param bytes Define the number of bytes required to represent the native seed.
     */
    NativeSeedType(Class<?> type, int bytes) {
        this.type = type;
        this.bytes = bytes;
    }

    /**
     * Gets the class type of the native seed.
     *
     * @return the type
     */
    public Class<?> getType() {
        return type;
    }

    /**
     * Gets the number of bytes required to represent the native seed type. If the type is
    // … 省略 141 行
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-123  ·  MethodSource

**项目** `JMeter`  **文件** `jmeter/src/dist-check/src/test/java/org/apache/jmeter/junit/JMeterTest.java`  **测试** `elementShouldNotBeModifiedWithConfigureModify`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource("guiComponents")
    public void elementShouldNotBeModifiedWithConfigureModify(GuiComponentHolder componentHolder) {
        JMeterGUIComponent guiItem = componentHolder.getComponent();
        TestElement expected = guiItem.createTestElement();
        TestElement actual = guiItem.createTestElement();
        guiItem.configure(actual);
        if (!Objects.equals(expected, actual)) {
            boolean breakpointForDebugging = Objects.equals(expected, actual);
            String expectedStr = new DslPrinterTraverser(DslPrinterTraverser.DetailLevel.ALL).append(expected).toString();
            String actualStr = new DslPrinterTraverser(DslPrinterTraverser.DetailLevel.ALL).append(actual).toString();
            assertEquals(
                    expectedStr,
                    actualStr,
                    () -> "TestElement should not be modified by " + guiItem.getClass().getName() + ".configure(element)"
            );
        }
        guiItem.modifyTestElement(actual);
        if (guiItem.getClass() == GraphQLHTTPSamplerGui.class) {
            // GraphQL sampler computes its arguments, so we don't compare them
            // See org.apache.jmeter.protocol.http.config.gui.GraphQLUrlConfigGui.modifyTestElement
            expected.removeProperty(HTTPSamplerBaseSchema.INSTANCE.getArguments());
            actual.removeProperty(HTTPSamplerBaseSchema.INSTANCE.getArguments());
        }
        if (!Objects.equals(expected, actual)) {
            if (improperlyUsesUiPlaceholders(guiItem.getClass())) {
                return;
            }
            boolean breakpointForDebugging = Objects.equals(expected, actual);
            String expectedStr = new DslPrinterTraverser(DslPrinterTraverser.DetailLevel.ALL).append(expected).toString();
            String actualStr = new DslPrinterTraverser(DslPrinterTraverser.DetailLevel.ALL).append(actual).toString();
            assertEquals(
                    expectedStr,
                    actualStr,
                    () -> "TestElement should not be modified by " + guiItem.getClass().getName() + ".configure(element); gui.modifyTestElement(element)"
            );
        }
    }
```

### Parameter provider — 同文件内的 `guiComponents`

```java
    static Collection<GuiComponentHolder> guiComponents() throws Throwable {
        List<GuiComponentHolder> components = new ArrayList<>(customGuiComponents());
        for (Object o : getObjects(TestBean.class)) {
            Class<?> c = o.getClass();
            JMeterGUIComponent item = new TestBeanGUI(c);
            components.add(new GuiComponentHolder(item));
        }
        return components;
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-126  ·  EnumSource

**项目** `Log4j`  **文件** `logging-log4j2/log4j-core-test/src/test/java/org/apache/logging/log4j/core/async/AsyncThreadContextGarbageFreeTest.java`  **测试** `testAsyncLogWritesToLog`

### 测试方法

```java
    @ParameterizedTest
    @EnumSource
    void testAsyncLogWritesToLog(final Mode asyncMode) throws Exception {
        testAsyncLogWritesToLog(ContextImpl.GARBAGE_FREE, asyncMode, loggingPath);
    }
```

### 枚举声明 — `Mode`（logging-log4j2/log4j-core-test/src/test/java/org/apache/logging/log4j/core/async/AbstractAsyncThreadContextTestBase.java）

```java
    protected enum Mode {
        ALL_ASYNC,
        MIXED,
        BOTH_ALL_ASYNC_AND_MIXED;

        void initSelector() {
            final ContextSelector selector;
            if (this == ALL_ASYNC || this == BOTH_ALL_ASYNC_AND_MIXED) {
                selector = new AsyncLoggerContextSelector();
            } else {
                selector = new ClassLoaderContextSelector();
            }
            LogManager.setFactory(new Log4jContextFactory(selector));
        }

        void initConfigFile() {
            // NOTICE: PLEASE DON'T REFACTOR: keep "file" local variable for confirmation in debugger.
            final String file = this == ALL_ASYNC //
                    ? "AsyncLoggerThreadContextTest.xml" //
                    : "AsyncLoggerConfigThreadContextTest.xml";
            props.setProperty(ConfigurationFactory.CONFIGURATION_FILE_PROPERTY, file);
        }
    }
```

### 请判定

`equivalence_class` / `enum_representation` / `enum_exploitation` / `behavior_carrying`

---

## IRR-129  ·  MethodSource

**项目** `Commons-RNG`  **文件** `commons-rng/commons-rng-core/src/test/java/org/apache/commons/rng/core/SplittableProvidersParametricTest.java`  **测试** `testSplitsMethodsUseSameSpliterator`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource("getSplittableProviders")
    void testSplitsMethodsUseSameSpliterator(SplittableUniformRandomProvider generator) {
        final long size = 10;
        final Spliterator<SplittableUniformRandomProvider> s = generator.splits(size, generator).spliterator();
        Assertions.assertEquals(s.getClass(), generator.splits().spliterator().getClass());
        Assertions.assertEquals(s.getClass(), generator.splits(size).spliterator().getClass());
        Assertions.assertEquals(s.getClass(), generator.splits(ThreadLocalGenerator.INSTANCE).spliterator().getClass());
    }
```

### Parameter provider — 同文件内的 `getSplittableProviders`

```java
    private static Iterable<SplittableUniformRandomProvider> getSplittableProviders() {
        return ProvidersList.listSplittable();
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-132  ·  MethodSource

**项目** `Hive`  **文件** `hive/ql/src/test/org/apache/hadoop/hive/ql/io/parquet/serde/TestParquetTimestampsHive2Compatibility.java`  **测试** `testWriteHive4UsingLegacyConversionReadHive2`

### 测试方法

```java
  @ParameterizedTest(name = "{0}")
  @MethodSource("generateTimestamps")
  void testWriteHive4UsingLegacyConversionReadHive2(String timestampString) {
    NanoTime nt = writeHive4(timestampString, TimeZone.getDefault().getID(), true);
    java.sql.Timestamp ts = readHive2(nt);
    assertEquals(timestampString, ts.toString());
  }
```

### Parameter provider — 同文件内的 `generateTimestamps`

```java

  private static Stream<String> generateTimestamps() {
    return Stream.concat(Stream.generate(new Supplier<String>() {
      int i = 0;

      @Override
      public String get() {
        StringBuilder sb = new StringBuilder(29);
        int year = (i % 9999) + 1;
        sb.append(zeros(4 - digits(year)));
        sb.append(year);
        sb.append('-');
        int month = (i % 12) + 1;
        sb.append(zeros(2 - digits(month)));
        sb.append(month);
        sb.append('-');
        int day = (i % 28) + 1;
        sb.append(zeros(2 - digits(day)));
        sb.append(day);
        sb.append(' ');
        int hour = i % 24;
        sb.append(zeros(2 - digits(hour)));
        sb.append(hour);
        sb.append(':');
        int minute = i % 60;
        sb.append(zeros(2 - digits(minute)));
        sb.append(minute);
        sb.append(':');
        int second = i % 60;
        sb.append(zeros(2 - digits(second)));
        sb.append(second);
        sb.append('.');
        // Bitwise OR with one to avoid times with trailing zeros
        int nano = (i % 1000000000) | 1;
        sb.append(zeros(9 - digits(nano)));
        sb.append(nano);
        i++;
        return sb.toString();
      }
    })
    // Exclude dates falling in the default Gregorian change date since legacy code does not handle that interval
    // gracefully. It is expected that these do not work well when legacy APIs are in use. 
    .filter(s -> !s.startsWith("1582-10"))
    .limit(3000), Stream.of("9999-12-31 23:59:59.999"));
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-135  ·  ValueSource

**项目** `Commons-HttpClient`  **文件** `httpcomponents-client/httpclient5-testing/src/test/java/org/apache/hc/client5/testing/async/TestConnectionClosureRace.java`  **测试** `testSpacedOutBatchesOfRequests`

### 测试方法

```java
    @ParameterizedTest(name = "Validation: {0}")
    @ValueSource(booleans = { false, true })
    @Timeout(5)
    @Order(5)
    void testSpacedOutBatchesOfRequests(final boolean validateConnections) throws Exception {
        try (final CloseableHttpAsyncClient client = asyncClient(validateConnections)) {
            final List<Future<SimpleHttpResponse>> futures = new ArrayList<>();
            for (int i = 0; i < 5; i++) {
                final List<Future<SimpleHttpResponse>> batchFutures = sendRequestBatch(client, 3);
                futures.addAll(batchFutures);
            }

            checkResults(getValidationPrefix(validateConnections) + "Multiple small batches", futures);
        }
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-138  ·  ValueSource

**项目** `Commons-Compress`  **文件** `commons-compress/src/test/java/org/apache/commons/compress/harmony/pack200/NewAttributeBandsTest.java`  **测试** `testIntegralLayouts`

### 测试方法

```java
    @ParameterizedTest
    @ValueSource(strings = { "B", "FB", "SB", "H", "FH", "SH", "I", "FI", "SI", "PB", "OB", "OSB", "POB", "PH", "OH", "OSH", "POH", "PI", "OI", "OSI", "POI" })
    void testIntegralLayouts(final String layoutStr) throws IOException {
        final CPUTF8 name = new CPUTF8("TestAttribute");
        final CPUTF8 layout = new CPUTF8(layoutStr);
        final MockNewAttributeBands newAttributeBands = new MockNewAttributeBands(1, null, null,
                new AttributeDefinition(35, AttributeDefinitionBands.CONTEXT_CLASS, name, layout));
        final List<AttributeLayoutElement> layoutElements = newAttributeBands.getLayoutElements();
        assertEquals(1, layoutElements.size());
        final Integral element = (Integral) layoutElements.get(0);
        assertEquals(layoutStr, element.getTag());
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-141  ·  MethodSource

**项目** `Commons-Numbers`  **文件** `commons-numbers/commons-numbers-arrays/src/test/java/org/apache/commons/numbers/arrays/SelectionTest.java`  **测试** `testIntDualPivotQuickSelectMaxRecursion`

### 测试方法

```java
    @ParameterizedTest
    @MethodSource(value = {"testIntPartition", "testIntPartitionBigData"})
    void testIntDualPivotQuickSelectMaxRecursion(int[] values, int[] indices) {
        assertPartition(values, indices, (a, k, n) -> {
            final int right = a.length - 1;
            if (right < 1 || k.length == 0) {
                return;
            }
            QuickSelect.dualPivotQuickSelect(a, 0, right,
                IndexSupport.createUpdatingInterval(k, k.length),
                QuickSelect.dualPivotFlags(2, 5));
        }, false);
    }
```

### Parameter provider — 同文件内的 `testIntPartition`

```java

    static Stream<Arguments> testIntPartition() {
        final Stream.Builder<Arguments> builder = Stream.builder();
        UniformRandomProvider rng = RandomSource.XO_SHI_RO_128_PP.create(123);
        // Sizes above and below the threshold for partitioning.
        // The largest size should trigger single-pivot sub-sampling for pivot selection.
        for (final int size : new int[] {5, 47, SU + 10}) {
            final int halfSize = size >>> 1;
            final int from = -halfSize;
            final int to = -halfSize + size;
            final int[] values = IntStream.range(from, to).toArray();
            final int[] zeros = values.clone();
            final int quarterSize = size >>> 2;
            Arrays.fill(zeros, quarterSize, halfSize + quarterSize, 0);
            for (final int k : new int[] {1, 2, 3, size}) {
                for (int i = 0; i < 15; i++) {
                    // Note: Duplicate indices do not matter
                    final int[] indices = rng.ints(k, 0, size).toArray();
                    builder.add(Arguments.of(
                        ArraySampler.shuffle(rng, values.clone()),
                        indices.clone()));
                    builder.add(Arguments.of(
                        ArraySampler.shuffle(rng, zeros.clone()),
                        indices.clone()));
                }
            }
            // Test sequential processing by creating potential ranges
            // after an initial low point. This should be high enough
            // so any range analysis that joins indices will leave the initial
            // index as a single point.
            final int limit = 50;
            if (size > limit) {
                for (int i = 0; i < 10; i++) {
                    final int[] indices = rng.ints(size - limit, limit, size).toArray();
                    // This sets a low index
                    indices[rng.nextInt(indices.length)] = rng.nextInt(0, limit >>> 1);
                    builder.add(Arguments.of(
                        ArraySampler.shuffle(rng, values.clone()),
                        indices.clone()));
                }
            }
            // min; max; min/max
            builder.add(Arguments.of(values.clone(), new int[] {0}));
            builder.add(Arguments.of(values.clone(), new int[] {size - 1}));
            builder.add(Arguments.of(values.clone(), new int[] {0, size - 1}));
            builder.add(Arguments.of(zeros.clone(), new int[] {0}));
            builder.add(Arguments.of(zeros.clone(), new int[] {size - 1}));
            builder.add(Arguments.of(zeros.clone(), new int[] {0, size - 1}));
        }
        final int value = Integer.MIN_VALUE;
        builder.add(Arguments.of(new int[] {}, new int[0]));
        builder.add(Arguments.of(new int[] {value}, new int[] {0}));
        builder.add(Arguments.of(new int[] {0, value}, new int[] {1}));
        builder.add(Arguments.of(new int[] {value, value, value}, new int[] {2}));
        builder.add(Arguments.of(new int[] {value, 0, 0, value}, new int[] {3}));
        builder.add(Arguments.of(new int[] {value, 0, 0, value}, new int[] {1, 2}));
        builder.add(Arguments.of(new int[] {value, 0, 1, 0, value}, new int[] {1, 3}));
        builder.add(Arguments.of(new int[] {value, 0, 0}, new int[] {0, 2}));
        builder.add(Arguments.of(new int[] {value, 123, 0, -456, 0, value}, new int[] {0, 1, 3}));
        // Dual-pivot with a large middle region (> 5 / 8) requires equal elements loop
        final int n = 128;
        final int[] x = IntStream.range(0, n).toArray();
        // Put equal elements in the central region:
        //          2/16      6/16             10/16      14/16
        // |  <P1    |    P1   |   P1< & < P2    |    P2    |    >P2    |
        final int sixteenth = n / 16;
        final int i2 = 2 * sixteenth;
        final int i6 = 6 * sixteenth;
        final int p1 = x[i2];
        final int p2 = x[n - i2];
        // Lots of values equal to the pivots
        Arrays.fill(x, i2, i6, p1);
        Arrays.fill(x, n - i6, n - i2, p2);
        // Equal value in between the pivots
        Arrays.fill(x, i6, n - i6, (p1 + p2) / 2);
        // ArraySampler.shuffle this and partition in the middle.
        // Also partition with the pivots in P1 and P2 using thirds.
        final int third = (int) (n / 3.0);
        // Use a fix seed to ensure we hit coverage with only 5 loops.
        rng = RandomSource.XO_SHI_RO_128_PP.create(-8111061151820577011L);
        for (int i = 0; i < 5; i++) {
            builder.add(Arguments.of(ArraySampler.shuffle(rng, x.clone()), new int[] {n >> 1}));
            builder.add(Arguments.of(ArraySampler.shuffle(rng, x.clone()),
                new int[] {third, 2 * third}));
        }
        // A single value smaller/greater than the pivot at the left/right/both ends
        Arrays.fill(x, 1);
        for (int i = 0; i <= 2; i++) {
            for (int j = 0; j <= 2; j++) {
                x[n - 1] = i;
                x[0] = j;
                builder.add(Arguments.of(x.clone(), new int[] {50}));
            }
        }
        // Reverse data. Makes it simple to detect failed range selection.
        final int[] a = IntStream.range(0, 50).toArray();
        for (int i = -1, j = a.length; ++i < --j;) {
            final int v = a[i];
            a[i] = a[j];
            a[j] = v;
        }
        builder.add(Arguments.of(a, new int[] {1, 1}));
        builder.add(Arguments.of(a, new int[] {1, 2}));
        builder.add(Arguments.of(a, new int[] {10, 12}));
        builder.add(Arguments.of(a, new int[] {10, 42}));
        builder.add(Arguments.of(a, new int[] {1, 48}));
        builder.add(Arguments.of(a, new int[] {48, 49}));
        return builder.build();
    }
```

### Parameter provider — 同文件内的 `testIntPartitionBigData`

```java

    static Stream<Arguments> testIntPartitionBigData() {
        final Stream.Builder<Arguments> builder = Stream.builder();
        final UniformRandomProvider rng = RandomSource.XO_SHI_RO_128_PP.create(123);
        // Sizes above the threshold (1200) for recursive partitioning
        for (final int size : new int[] {1000, 5000, 10000}) {
            final int[] a = IntStream.range(0, size).toArray();
            // With repeat elements
            final int[] b = rng.ints(size, 0, size >> 3).toArray();
            for (int i = 0; i < 15; i++) {
                builder.add(Arguments.of(
                    ArraySampler.shuffle(rng, a.clone()),
                    new int[] {rng.nextInt(size)}));
                builder.add(Arguments.of(b.clone(),
                    new int[] {rng.nextInt(size)}));
            }
        }
        // Hit Floyd-Rivest sub-sampling conditions.
        // Close to edge but outside edge select size.
        final int n = 7000;
        final int[] x = IntStream.range(0, n).toArray();
        builder.add(Arguments.of(x.clone(), new int[] {20}));
        builder.add(Arguments.of(x.clone(), new int[] {n - 1 - 20}));
        // Constant value when using FR partitioning
        Arrays.fill(x, 123);
        builder.add(Arguments.of(x, new int[] {x.length >>> 1}));
        return builder.build();
    }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---

## IRR-144  ·  CsvSource

**项目** `JMeter`  **文件** `jmeter/src/protocol/http/src/test/java/org/apache/jmeter/protocol/http/proxy/DefaultSamplerCreatorTest.java`  **测试** `computeSamplerNameWithCounter`

### 测试方法

```java
    void computeSamplerNameWithCounter(int sampleNameMode, String format, String expectedName) {
        DefaultSamplerCreator samplerCreator = new DefaultSamplerCreator();
        HTTPSamplerBase sampler = new HTTPSampler();
        sampler.setPath("/some/path");
        sampler.setDomain("jmeter.invalid");
        sampler.setMethod("GET");
        sampler.setPort(443);
        sampler.setProtocol("https");
        HttpRequestHdr request = new HttpRequestHdr(
                "prefix|",
                "samplerName",
                sampleNameMode,
                format
        );
        samplerCreator.setCounter(41);
        samplerCreator.computeSamplerName(sampler, request);
        assertEquals(expectedName, sampler.getName());
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-147  ·  CsvSource

**项目** `JMeter`  **文件** `jmeter/src/components/src/test/java/org/apache/jmeter/assertions/TestJSONPathAssertion.java`  **测试** `testGetResult_pathsWithOneResult`

### 测试方法

```java
    void testGetResult_pathsWithOneResult(String data, String jsonPath, String expectedResult) {
        SampleResult samplerResult = new SampleResult();
        samplerResult.setResponseData(data.getBytes(Charset.defaultCharset()));

        JSONPathAssertion instance = new JSONPathAssertion();
        instance.setJsonPath(jsonPath);
        instance.setJsonValidationBool(true);
        instance.setExpectedValue(expectedResult);
        AssertionResult expResult = new AssertionResult("");
        AssertionResult result = instance.getResult(samplerResult);
        assertEquals(expResult.getName(), result.getName());
        assertFalse(result.isFailure());
    }
```

### 请判定

`equivalence_class` / `semantic_role`

---

## IRR-150  ·  MethodSource

**项目** `Avro`  **文件** `avro/lang/java/avro/src/test/java/org/apache/avro/io/TestResolvingIO.java`  **测试** `testIdentical`

### 测试方法

```java
  @ParameterizedTest
  @MethodSource("data2")
  public void testIdentical(Encoding encoding, int skip, String jsonWriterSchema, String writerCalls,
      String jsonReaderSchema, String readerCalls) throws IOException {
    performTest(encoding, skip, jsonWriterSchema, writerCalls, jsonWriterSchema, writerCalls);
  }
```

### Parameter provider — 同文件内的 `data2`

```java

  public static Stream<Arguments> data2() {
    return TestValidatingIO.convertTo2dStream(encodings, skipLevels, testSchemas());
  }
```

### 请判定

`equivalence_class` / `semantic_role` / `methodsource_intent` / `value_complexity` / `behavior_carrying`

---
