Update Note:

2025/9/1
pid1_config_check.py
重點改動

1. 全新「制式解析流程」
    由 model.ini 的 tvSysMap 進入 → 讀取 tvSysMapCfgs.xml → 擷取 <CountryTvSysMapXML> → 逐檔開啟 countryTvSysMap.xml，在每個

<COUNTRY_TVCONFIG_MAP>
  <COUNTRY_NAME>...</COUNTRY_NAME>
  <TV_SYSTEM>...</TV_SYSTEM>
  <TV_CONFIG>...</TV_CONFIG>
</COUNTRY_TVCONFIG_MAP>

解析出 COUNTRY_NAME → TV_SYSTEM，並轉成 ISO 兩碼（alpha-2）國碼。

2. EU/EFTA/GB/CH 合規檢查（以 countryTvSysMap 為準）
    對屬於 EU + EFTA + GB + CH 的國家，要求其 TV_SYSTEM 必須是 DVB 或 DVB_CO。不符合者會被標為違規並列出清單。

3. 國名對應與解析更健壯

    3.1. 新增 國名→兩碼 映射（如 UNITED_KINGDOM→GB, SOUTH_KOREA→KR），同時仍支援原本兩碼國碼。
    3.2. XML 解析為主、regex 為備援；可容忍縮排與空白差異、支援多個 <CountryTvSysMapXML> 檔的彙整。

4. 保留並強化既有檢查

    4.1. tvSysMap → [VOLUME_CURVE_CFG]：value 中路徑解析與存在性檢查。
    4.2 <TvSystem type="DVB|DVB_CO|DTMB"> 是否存在；其區塊內 inputSource 的 DVBT/DVBC/DVBS 至少一個非 NULL。
    4.3 CLTV（Live TV）與 多制式切換 欄位資訊展示（不影響通過與否，除非值無法解析）。

執行流程（High level）

1. 讀取 sys/*.ini（可選加上 device_sys.ini 覆寫）→ 取得 Model_1/Board_1 路徑與存在性。
2. 從 model.ini 解析 tvSysMap → 開啟 tvSysMapCfgs.xml。
3. 抽取 <CountryTvSysMapXML> 清單 → 逐檔開啟 countryTvSysMap.xml：
    3.1. 解析 <COUNTRY_NAME>、<TV_SYSTEM>（→ 建立 alpha2:TV_SYSTEM 對照）。
    3.2. 生成 countries（兩碼集合）與 ctvs_map（兩碼→制式）。
    3.3. 若本步驟未取得任何國家，再回退到舊法（以 model.ini 與其引用檔掃描兩碼字樣）。
4. EU 合規：對 countries ∩ (EU+EFTA+GB+CH)，檢查對應 TV_SYSTEM ∈ {DVB, DVB_CO}；否則記為違規。
5. TvSystem/InputSource 檢查（獨立於上一步）：
    5.1. 驗證至少存在一個 <TvSystem type="DVB|DVB_CO|DTMB">；
    5.2. 蒐集其 inputSource 值，確認 DVBT/DVBC/DVBS 至少一個非 NULL。
6. VOLUME_CURVE_CFG 檔案：解析並確認所有引用路徑存在。
7. 報表與退出碼：彙總狀態、違規原因；如指定 --fail-warning 且有錯誤則以非 0 結束。

2025/8/28
=======
共有13種類別可以使用python判斷，可以檢視90%以上的場景，其他device的特例需要另外增加

<img width="1767" height="944" alt="image" src="https://github.com/user-attachments/assets/83745e33-bf55-4f03-9017-39609adef47f" />
<img width="1756" height="538" alt="image" src="https://github.com/user-attachments/assets/295e65af-6d89-4c35-b440-5ff0cb4d1764" />
