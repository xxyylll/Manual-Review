# Parameterized Test Manual Review

这个仓库包含本次 Parameterized Test Study 的人工 review 材料。

## Review 流程

1. 请先阅读 `rater_guide.md`，了解每个分类维度的定义和 allowed values。
2. 然后打开分配给你的 review packet。
3. 每个 test 最后都有一个 **To classify** 部分，请只填写这里列出的分类项。
4. 在对应的 rating sheet 中，根据相同的 `item_id` 找到对应行并填写。
5. 如果某一分类不适用于当前 test/source，请保持该单元格为空。
6. 如果代码中的信息不足以做出判断，可以选择 `unclear`，并在 `notes` 中简单说明原因。
7. 请为每个 case 填写 `confidence`。

所有需要判断的代码都已经包含在 packet 中，不需要下载、编译或运行原项目。

## Reviewer Assignment

请只 review 分配给你自己的 packet，并独立完成判断，不要参考其他 reviewer 的结果。

- Reviewer 1 → `[ReviewerName1]_review_packet.md`
- Reviewer 2 → `[ReviewerName2]_review_packet.md`
- Reviewer 3 → `[ReviewerName3]_review_packet.md`

## Files

- `rater_guide.md` — 分类规则与 codebook
- `*_review_packet.md` — 每位 reviewer 对应的 test cases

最终结果请填写到单独提供的 rating sheet 中。
