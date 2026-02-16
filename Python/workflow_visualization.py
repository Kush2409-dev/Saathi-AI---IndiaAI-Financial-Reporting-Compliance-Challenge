"""
Workflow Visualization

Generate visual diagrams of the audit workflow
"""


def print_workflow_diagram():
    """Print ASCII diagram of the workflow"""
    diagram = """
    ╔══════════════════════════════════════════════════════════════╗
    ║           AUDIT COMPLIANCE AGENTIC WORKFLOW                  ║
    ║                    (LangGraph Based)                         ║
    ╚══════════════════════════════════════════════════════════════╝
    
    
         ┌─────────────────────────────────────────┐
         │  INPUT: Company Documents + Audit Rules │
         └──────────────────┬──────────────────────┘
                            │
                            ▼
         ┌──────────────────────────────────────────┐
         │   NODE 1: Extract Audit Questions        │
         │   • Parse audit configuration            │
         │   • Load rule-specific questions         │
         │   • Initialize state                     │
         └──────────────────┬───────────────────────┘
                            │
                            ▼
    ┌────────────────────────────────────────────────────────┐
    │                  PROCESSING LOOP                        │
    │  (Iterates through each audit question sequentially)   │
    └────────────────────────────────────────────────────────┘
                            │
                            ▼
         ┌──────────────────────────────────────────┐
         │   NODE 2: Process Single Question        │
         │   • Read current question                │
         │   • Access previous responses (context)  │
         │   • Analyze documents                    │
         │   • Apply LLM reasoning                  │
         │   • Extract evidence with page refs      │
         └──────────────────┬───────────────────────┘
                            │
                            ▼
         ┌──────────────────────────────────────────┐
         │   NODE 3: Validate Response              │
         │   • Quality check the answer             │
         │   • Verify evidence relevance            │
         │   • Assess reasoning logic               │
         │   • Confirm confidence score             │
         └──────────────────┬───────────────────────┘
                            │
                            ▼
         ┌──────────────────────────────────────────┐
         │   NODE 4: Check Failure Triggers         │
         │   • Evaluate compliance conditions       │
         │   • Check if failure criteria met        │
         │   • Update compliance status             │
         └──────────────────┬───────────────────────┘
                            │
                            ▼
         ┌──────────────────────────────────────────┐
         │   NODE 5: Increment Question Index       │
         │   • Move to next question                │
         └──────────────────┬───────────────────────┘
                            │
                            ▼
         ┌──────────────────────────────────────────┐
         │   ROUTING: More Questions?               │
         │   • Yes → Loop back to NODE 2            │
         │   • No  → Continue to final output       │
         └──────────────────┬───────────────────────┘
                            │ (All questions done)
                            ▼
         ┌──────────────────────────────────────────┐
         │   NODE 6: Generate Final Output          │
         │   • Aggregate all findings               │
         │   • Determine final compliance status    │
         │   • Create summary finding               │
         │   • Calculate confidence score           │
         │   • Check for auditor oversight          │
         └──────────────────┬───────────────────────┘
                            │
                            ▼
         ┌──────────────────────────────────────────┐
         │  OUTPUT: Structured Audit Report         │
         │  • Company Name                          │
         │  • Rule ID                               │
         │  • Compliance Status                     │
         │  • Summary Finding                       │
         │  • Auditor Oversight                     │
         │  • Reasoning Path                        │
         │  • Evidence Snippets                     │
         │  • Page References                       │
         │  • Confidence Score                      │
         └──────────────────────────────────────────┘


    ╔══════════════════════════════════════════════════════════════╗
    ║                    KEY FEATURES                              ║
    ╚══════════════════════════════════════════════════════════════╝
    
    1. SEQUENTIAL PROCESSING
       → Each question builds on previous answers
       → Context flows through the workflow
       → Dependencies are automatically managed
    
    2. VALIDATION AT EACH STEP
       → Senior auditor review simulation
       → Quality assurance built-in
       → Confidence scoring
    
    3. FAILURE TRIGGER DETECTION
       → Real-time compliance checking
       → Early exit possible if critical issues found
       → Rule-based logic evaluation
    
    4. EVIDENCE TRACKING
       → Exact quotes from documents
       → Page number references
       → Searchable evidence trail
    
    5. FLEXIBLE ARCHITECTURE
       → Easy to add new audit rules
       → Extensible with custom nodes
       → Can integrate with external systems
    """
    
    print(diagram)


def generate_mermaid_diagram() -> str:
    """Generate Mermaid diagram for documentation"""
    mermaid = """```mermaid
graph TD
    A[Input: Documents + Config] --> B[Extract Audit Questions]
    B --> C{For Each Question}
    C --> D[Process Question with LLM]
    D --> E[Validate Response]
    E --> F[Check Failure Triggers]
    F --> G[Increment to Next Question]
    G --> C
    C -->|All Done| H[Generate Final Output]
    H --> I[Structured Audit Report]
    
    style A fill:#e1f5ff
    style I fill:#c8e6c9
    style D fill:#fff9c4
    style E fill:#ffe0b2
    style F fill:#ffccbc
    style H fill:#f8bbd0
```"""
    
    return mermaid


def print_state_flow():
    """Show how state flows through the workflow"""
    flow = """
    ╔══════════════════════════════════════════════════════════════╗
    ║                 STATE FLOW DIAGRAM                           ║
    ╚══════════════════════════════════════════════════════════════╝
    
    INITIAL STATE:
    {
      company_name: "Vikas WSP Limited",
      rule_id: "SEC_186_4_PURPOSE",
      audit_questions: [...],
      current_question_index: 0,
      question_responses: [],
      documents: [...]
    }
    
                    ↓ ↓ ↓
    
    AFTER QUESTION 1:
    {
      ...
      current_question_index: 1,
      question_responses: [
        {
          answer: "Yes",
          evidence: "CARO report mentions...",
          page_ref: 12,
          reasoning: "The auditor's CARO...",
          confidence: 95
        }
      ],
      evidence_snippets: ["CARO report mentions..."],
      page_refs: [12]
    }
    
                    ↓ ↓ ↓
    
    AFTER QUESTION 2:
    {
      ...
      current_question_index: 2,
      question_responses: [
        {...},  # Question 1 response
        {
          answer: "Yes - 120 days",
          evidence: "Default period stated as...",
          page_ref: 12,
          reasoning: "Building on Q1, the period...",
          confidence: 92
        }
      ],
      evidence_snippets: ["...", "Default period stated as..."],
      page_refs: [12, 12]
    }
    
                    ↓ ↓ ↓
    
    ... (continues for all questions) ...
    
                    ↓ ↓ ↓
    
    FINAL STATE:
    {
      ...
      compliance_status: "Non-Compliant",
      summary_finding: "The company has defaulted...",
      auditor_oversight: "Yes - The statutory auditor...",
      confidence_score: 92,
      question_responses: [Q1, Q2, Q3, Q4, Q5, Q6]
    }
    """
    
    print(flow)


def print_tool_integration_guide():
    """Show how to add custom tools/validators"""
    guide = """
    ╔══════════════════════════════════════════════════════════════╗
    ║          ADDING CUSTOM TOOLS & VALIDATORS                    ║
    ╚══════════════════════════════════════════════════════════════╝
    
    1. ADD CUSTOM VALIDATION TOOLS
    ────────────────────────────────────────────────────────────
    
    def custom_financial_validator(state: AuditState) -> AuditState:
        \"\"\"Validate financial calculations\"\"\"
        
        # Extract financial data
        finance_cost = extract_value(state, "finance_cost")
        borrowings = extract_value(state, "borrowings")
        
        # Calculate ratio
        ratio = (finance_cost / borrowings) * 100
        
        # Validate against threshold
        if ratio < 6.0:
            state["validation_flags"].append({
                "type": "financial_red_flag",
                "message": f"Interest rate {ratio}% below 6% threshold",
                "severity": "high"
            })
        
        return state
    
    # Add to workflow
    workflow.add_node("financial_validation", custom_financial_validator)
    workflow.add_edge("check_triggers", "financial_validation")
    
    
    2. ADD DOCUMENT CROSS-REFERENCE TOOL
    ────────────────────────────────────────────────────────────
    
    def cross_reference_checker(state: AuditState) -> AuditState:
        \"\"\"Check consistency across documents\"\"\"
        
        # Extract claims from different documents
        audit_report_claims = extract_from_doc(state, "audit_report")
        financial_stmt_claims = extract_from_doc(state, "financial_statements")
        
        # Compare for inconsistencies
        inconsistencies = find_inconsistencies(
            audit_report_claims,
            financial_stmt_claims
        )
        
        if inconsistencies:
            state["auditor_oversight"] = "Yes - Inconsistencies found"
            state["inconsistencies"] = inconsistencies
        
        return state
    
    
    3. ADD REGULATORY COMPLIANCE CHECKER
    ────────────────────────────────────────────────────────────
    
    def regulatory_checker(state: AuditState) -> AuditState:
        \"\"\"Check against specific regulations\"\"\"
        
        regulations = {
            "IND_AS_109": check_ind_as_109_compliance,
            "RBI_NORMS": check_rbi_compliance,
            "SECTION_186": check_section_186_compliance
        }
        
        for reg_name, checker_func in regulations.items():
            is_compliant = checker_func(state)
            state["regulatory_checks"][reg_name] = is_compliant
        
        return state
    
    
    4. INTEGRATE WITH EXTERNAL SYSTEMS
    ────────────────────────────────────────────────────────────
    
    def external_api_integration(state: AuditState) -> AuditState:
        \"\"\"Query external databases\"\"\"
        
        # Query company registry
        company_data = query_company_registry(state["company_name"])
        
        # Query credit rating agencies
        credit_rating = query_credit_rating(state["company_name"])
        
        # Enrich state with external data
        state["external_data"] = {
            "company_registry": company_data,
            "credit_rating": credit_rating
        }
        
        return state
    """
    
    print(guide)


if __name__ == "__main__":
    print_workflow_diagram()
    print("\n" + "="*70 + "\n")
    print_state_flow()
    print("\n" + "="*70 + "\n")
    print_tool_integration_guide()
    print("\n" + "="*70 + "\n")
    print("Mermaid Diagram Code:")
    print(generate_mermaid_diagram())
