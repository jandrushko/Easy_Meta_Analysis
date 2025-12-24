###############################################################################
# META-ANALYSIS SHINY APP
# 
#  Author: Justin Andrushko
###############################################################################

library(shiny)
library(shinydashboard)
library(shinyWidgets)
library(DT)
library(metafor)
library(clubSandwich)
library(dplyr)
library(tidyr)
library(ggplot2)
library(ggtext)
library(readxl)
library(plotly)
library(viridis)
library(colourpicker)

# ========================= UI ================================================
ui <- dashboardPage(
  dashboardHeader(title = "Meta-Analysis Suite"),
  
  dashboardSidebar(
    sidebarMenu(
      menuItem("Data Import", tabName = "import", icon = icon("upload")),
      menuItem("Advanced Filtering", tabName = "filter", icon = icon("filter")),
      menuItem("Effect Sizes", tabName = "effects", icon = icon("calculator")),
      menuItem("Meta-Analysis", tabName = "forest", icon = icon("chart-line")),
      menuItem("Funnel Plot", tabName = "funnel", icon = icon("filter")),
      menuItem("Export", tabName = "export", icon = icon("download"))
    )
  ),
  
  dashboardBody(
    tags$head(
      tags$style(HTML("
        .filter-box {
          background: #f8f9fa;
          padding: 15px;
          margin: 10px 0;
          border-radius: 5px;
          border: 1px solid #ddd;
        }
        .btn-group-xs > .btn {
          padding: 1px 5px;
          font-size: 11px;
        }
        .missing-data-warning {
          background: #fff3cd;
          border: 1px solid #ffc107;
          padding: 15px;
          border-radius: 5px;
          margin: 10px 0;
        }
        .summary-stats-box {
          background: #e7f3ff;
          border: 2px solid #0066cc;
          padding: 15px;
          border-radius: 5px;
          margin: 15px 0;
        }
        .control-section {
          background: #f5f5f5;
          padding: 12px;
          margin: 10px 0;
          border-radius: 5px;
          border: 1px solid #ddd;
        }
        .info-box-dependency {
          background: #e8f5e9;
          border: 1px solid #4caf50;
          padding: 10px;
          border-radius: 5px;
          margin-top: 10px;
          font-size: 11px;
        }
        .robust-results-box {
          background: #fff3e0;
          border: 2px solid #ff9800;
          padding: 15px;
          border-radius: 5px;
          margin: 10px 0;
        }
        .sensitivity-box {
          background: #e3f2fd;
          border: 2px solid #2196f3;
          padding: 15px;
          border-radius: 5px;
          margin: 10px 0;
        }
      "))
    ),
    
    tabItems(
      # =================== DATA IMPORT TAB =====================================
      tabItem(
        tabName = "import",
        fluidRow(
          box(
            title = "Data Import",
            status = "primary",
            solidHeader = TRUE,
            width = 12,
            
            fluidRow(
              column(6,
                fileInput("file", "Choose Excel/CSV File:",
                         accept = c(".xlsx", ".xls", ".csv")),
                uiOutput("sheet_selector"),
                actionButton("load_data", "Load Data", 
                           class = "btn-primary", icon = icon("upload"))
              ),
              column(6,
                h4("Import Summary"),
                verbatimTextOutput("import_summary")
              )
            ),
            
            hr(),
            
            h4("Data Preview"),
            DT::dataTableOutput("preview_table")
          )
        )
      ),
      
      # =================== ADVANCED FILTERING TAB ==============================
      tabItem(
        tabName = "filter",
        
        fluidRow(
          valueBoxOutput("n_total"),
          valueBoxOutput("n_filtered"),
          valueBoxOutput("n_excluded")
        ),
        
        fluidRow(
          box(
            title = "Filter Variables",
            status = "success",
            solidHeader = TRUE,
            width = 8,
            
            div(style = "background: #e7f3ff; padding: 10px; border-radius: 5px; margin-bottom: 15px;",
              h4("Add Filter Variables", style = "margin-top: 0;"),
              fluidRow(
                column(6,
                  selectInput("cat_var_to_add", "Categorical variables:",
                             choices = c())
                ),
                column(3,
                  br(),
                  actionButton("add_cat_var", "Add Variable",
                             class = "btn-info", icon = icon("plus"))
                ),
                column(3,
                  br(),
                  actionButton("remove_all_filters", "Remove All",
                             class = "btn-warning", icon = icon("times"))
                )
              )
            ),
            
            div(
              h4("Active Filters:"),
              uiOutput("active_filters_ui")
            ),
            
            hr(),
            
            actionButton("apply_filters", "Apply Filters", 
                       class = "btn-success btn-lg", icon = icon("check")),
            actionButton("reset_filters", "Reset to Original Data", 
                       class = "btn-warning", icon = icon("undo"))
          ),
          
          box(
            title = "Filter Summary",
            status = "info",
            solidHeader = TRUE,
            width = 4,
            
            h4("Current Status:"),
            verbatimTextOutput("filter_summary"),
            
            hr(),
            
            h4("Active Variables:"),
            verbatimTextOutput("active_vars_list")
          )
        ),
        
        fluidRow(
          box(
            title = "Filtered Data Preview",
            status = "primary",
            width = 12,
            DT::dataTableOutput("filtered_preview")
          )
        )
      ),
      
      # =================== EFFECT SIZES TAB ====================================
      tabItem(
        tabName = "effects",
        
        fluidRow(
          box(
            title = "Effect Size Configuration",
            status = "warning",
            solidHeader = TRUE,
            width = 12,
            
            fluidRow(
              column(3,
                h4("Effect Size Type"),
                selectInput("es_type", "Type:",
                           choices = c(
                             "SMCR (pre-post)" = "SMCR",
                             "SMD (between groups)" = "SMD",
                             "Raw difference (MD)" = "MD"
                           )),
                
                checkboxInput("apply_hedges", "Apply Hedges correction", 
                            value = TRUE),
                
                hr(),
                
                checkboxInput("auto_remove_missing", 
                            "Auto-remove rows with missing data", 
                            value = TRUE)
              ),
              
              column(4,
                h4("Column Mapping"),
                selectInput("studyid_col", "Study ID:", choices = c()),
                selectInput("mean1_col", "Mean 1 (Post):", choices = c()),
                selectInput("mean2_col", "Mean 2 (Pre):", choices = c()),
                selectInput("sd1_col", "SD 1:", choices = c()),
                selectInput("sd2_col", "SD 2:", choices = c()),
                selectInput("n_col", "Sample Size:", choices = c()),
                selectInput("group_col", "Group (optional):", choices = c("", ""))
              ),
              
              column(5,
                h4("Pre-Post Correlation"),
                radioButtons("corr_source", "Source:",
                           choices = c(
                             "Fixed value" = "fixed",
                             "From column" = "column",
                             "Compute from change SD" = "compute"
                           ),
                           selected = "column"),
                
                conditionalPanel(
                  condition = "input.corr_source == 'fixed'",
                  sliderInput("corr_value", "Correlation:", 
                            min = 0, max = 1, value = 0.5, step = 0.05)
                ),
                
                conditionalPanel(
                  condition = "input.corr_source == 'column'",
                  selectInput("corr_col", "Correlation column:", choices = c())
                ),
                
                conditionalPanel(
                  condition = "input.corr_source == 'compute'",
                  selectInput("change_sd_col", "Change SD column:", choices = c())
                )
              )
            ),
            
            hr(),
            
            fluidRow(
              column(12,
                div(style = "background: #fff3cd; padding: 10px; border-radius: 5px;",
                  strong("Pre-calculation Check:"),
                  verbatimTextOutput("precalc_check")
                )
              )
            ),
            
            uiOutput("missing_data_warning"),
            
            actionButton("calculate_es", "Calculate Effect Sizes",
                       class = "btn-warning btn-lg", icon = icon("calculator"))
          )
        ),
        
        fluidRow(
          box(
            title = "Effect Sizes",
            status = "primary",
            width = 8,
            DT::dataTableOutput("es_table")
          ),
          
          box(
            title = "Distribution",
            status = "info",
            width = 4,
            plotOutput("es_distribution", height = "300px"),
            hr(),
            verbatimTextOutput("es_summary")
          )
        )
      ),
      
      # =================== Meta-ANALYSIS TAB ==========================
      tabItem(
        tabName = "forest",
        
        # MODEL SETTINGS (formerly in Meta-Analysis tab)
        fluidRow(
          box(
            title = HTML("<i class='fa fa-cog'></i> Meta-Analysis Model Settings"),
            status = "warning",
            solidHeader = TRUE,
            width = 12,
            collapsible = TRUE,
            collapsed = FALSE,
            
            div(style = "background: #fff9e6; border: 2px solid #ffc107; padding: 12px; border-radius: 5px;",
              fluidRow(
                column(3,
                  h4("Model Type", style = "margin-top: 0;"),
                  selectInput("ma_model", "Estimation method:",
                             choices = c(
                               "Random Effects (REML)" = "REML",
                               "Fixed Effect" = "FE",
                               "DerSimonian-Laird" = "DL",
                               "Maximum Likelihood" = "ML"
                             ),
                             selected = "REML"),
                  helpText("REML is recommended for most meta-analyses")
                ),
                
                column(3,
                  h4("Statistical Test", style = "margin-top: 0;"),
                  selectInput("test_type", "Test type:",
                             choices = c(
                               "Knapp-Hartung" = "knha",
                               "z-test" = "z"
                             ),
                             selected = "knha"),
                  helpText("Knapp-Hartung adjusts for uncertainty in τ²")
                ),
                
                column(3,
                  h4("Multi-level Model", style = "margin-top: 0;"),
                  checkboxInput("use_multilevel", 
                              "Use 3-level model (StudyID/EffectID)",
                              value = FALSE),
                  helpText("Enable when multiple outcomes per study")
                ),
                
                column(3,
                  br(),
                  div(style = "background: #e3f2fd; padding: 10px; border-radius: 5px; margin-top: 15px;",
                    p(style = "margin: 0; font-size: 12px;",
                      icon("info-circle"),
                      " Settings apply when you click 'Update Plot' below.")
                  )
                )
              ),
              
              # === NEW: DEPENDENCY HANDLING OPTIONS ===
              hr(style = "border-color: #ffc107; margin: 15px 0;"),
              
              fluidRow(
                column(12,
                  h4(icon("link"), " Handling Dependent Effect Sizes", 
                     style = "margin-top: 0; margin-bottom: 10px; color: #856404;"),
                  p(style = "font-size: 12px; color: #666; margin-bottom: 15px;",
                    "When the same participants contribute multiple outcomes (e.g., strength AND size from the same 8 people), ",
                    "their effect sizes share sampling error. These options handle this dependency.")
                )
              ),
              
              fluidRow(
                column(4,
                  # CHE MODEL OPTION
                  checkboxInput("use_che_model", 
                              strong("Use CHE working model (vcalc)"),
                              value = FALSE),
                  
                  conditionalPanel(
                    condition = "input.use_che_model == true",
                    sliderInput("assumed_rho", 
                              "Assumed within-study correlation (ρ):",
                              min = 0, max = 0.95, value = 0.5, step = 0.05),
                    
                    div(class = "info-box-dependency",
                      icon("info-circle"), 
                      HTML("<strong>What is ρ (rho)?</strong><br>
                           The correlation between effect sizes from the same study/participants. 
                           When participants contribute to multiple outcomes, their sampling errors are correlated.<br><br>
                           <strong>Guidance:</strong><br>
                           • ρ = 0.5 is a reasonable default<br>
                           • Higher (0.6-0.8) for very similar outcomes<br>
                           • Lower (0.3-0.5) for different outcome types")
                    )
                  )
                ),
                
                column(4,
                  # ROBUST VARIANCE ESTIMATION OPTION
                  checkboxInput("use_robust_ve", 
                              strong("Use robust variance estimation"),
                              value = FALSE),
                  
                  conditionalPanel(
                    condition = "input.use_robust_ve == true",
                    
                    div(class = "info-box-dependency",
                      icon("shield-alt"), 
                      HTML("<strong>What is robust VE?</strong><br>
                           Cluster-robust standard errors (using clubSandwich) protect against 
                           misspecification of the dependency structure.<br><br>
                           <strong>Key benefit:</strong> Even if your assumed ρ is wrong, 
                           statistical inference remains valid.<br><br>
                           <em>Recommended when you have multiple effect sizes per study.</em>")
                    )
                  )
                ),
                
                column(4,
                  # SENSITIVITY ANALYSIS OPTION
                  checkboxInput("run_rho_sensitivity", 
                              strong("Run sensitivity analysis"),
                              value = FALSE),
                  
                  conditionalPanel(
                    condition = "input.run_rho_sensitivity == true",
                    
                    div(class = "info-box-dependency",
                      icon("chart-line"), 
                      HTML("<strong>What is sensitivity analysis?</strong><br>
                           Tests whether conclusions change with different assumed ρ values 
                           (0.0, 0.3, 0.5, 0.7, 0.9).<br><br>
                           <strong>Interpretation:</strong> If estimates and CIs are stable 
                           across ρ values, your findings are robust to the correlation assumption.")
                    )
                  )
                )
              )
            ),
            
            # RUN ANALYSIS BUTTON
            hr(style = "border-color: #ffc107; margin: 15px 0;"),
            div(style = "text-align: center;",
              actionButton("run_analysis", 
                         HTML("<i class='fa fa-play'></i> Run Meta-Analysis"),
                         class = "btn-warning btn-lg",
                         style = "font-size: 16px; padding: 12px 30px;"),
              p(style = "margin-top: 10px; color: #666; font-size: 12px;",
                "Click after configuring model settings above. Required for CHE model and sensitivity analysis.")
            )
          )
        ),
        
        # FOREST PLOT CONTROLS
        fluidRow(
          box(
            title = "Enhanced Forest Plot Controls",
            status = "success",
            solidHeader = TRUE,
            width = 12,
            collapsible = TRUE,
            
            fluidRow(
              column(3,
                div(class = "control-section",
                  h4("Data Type", style = "margin-top: 0;"),
                  radioButtons("forest_data_type", NULL,
                             choices = c(
                               "Arm-level effects" = "arm",
                               "Between-group differences" = "diff"
                             ),
                             selected = "arm"),
                  
                  # Conditional panel for between-group options
                  conditionalPanel(
                    condition = "input.forest_data_type == 'diff'",
                    hr(style = "margin: 8px 0;"),
                    
                    # Group selection
                    selectInput("diff_group1", "Treatment group:", choices = c()),
                    selectInput("diff_group2", "Control group:", choices = c()),
                    
                    # Treatment subgroup variable (e.g., Training_Type)
                    selectInput("diff_treatment_subgroup", "Split treatment by:",
                               choices = c("(none)" = ""),
                               selected = ""),
                    helpText(style = "font-size: 10px;", 
                            "e.g., Training_Type to show Isometric vs Control and Concentric vs Control separately"),
                    
                    hr(style = "margin: 8px 0;"),
                    
                    radioButtons("diff_aggregation", "Difference calculation:",
                               choices = c(
                                 "One per study (averaged)" = "averaged",
                                 "Multiple per study (by outcome)" = "multiple"
                               ),
                               selected = "averaged"),
                    
                    # Matching variables for multiple differences
                    conditionalPanel(
                      condition = "input.diff_aggregation == 'multiple'",
                      selectInput("diff_match_vars", "Match outcomes by:",
                                choices = c(),
                                multiple = TRUE),
                      helpText(style = "font-size: 10px;", 
                              "Select columns that define unique outcomes (e.g., Measure, Muscle_Group)")
                    )
                  )
                )
              ),
              
              column(3,
                div(class = "control-section",
                  h4("Visual Style", style = "margin-top: 0;"),
                  selectInput("forest_style", "Style:",
                           choices = c(
                             "Classic" = "classic",
                             "With distributions" = "distribution"
                           )),
                  checkboxInput("show_ci_band", "Show CI band", value = TRUE),
                  checkboxInput("show_summary_stats", "Show summary stats below", value = TRUE)
                )
              ),
              
              column(3,
                div(class = "control-section",
                  h4("Distributions", style = "margin-top: 0;"),
                  conditionalPanel(
                    condition = "input.forest_style == 'distribution'",
                    checkboxInput("show_shaded_dist", "Shaded areas", value = TRUE),
                    sliderInput("dist_transparency", "Transparency:",
                              min = 0.1, max = 1, value = 0.4, step = 0.05),
                    sliderInput("dist_width", "Width:",
                              min = 0.2, max = 1.5, value = 0.8, step = 0.1)
                  )
                )
              ),
              
              column(3,
                div(class = "control-section",
                  h4("Appearance", style = "margin-top: 0;"),
                  colourInput("forest_color", "Line color:", value = "#1b9e77"),
                  colourInput("forest_fill_color", "Fill color:", value = "#1b9e77"),
                  sliderInput("row_spacing", "Row spacing:",
                            min = 0.3, max = 2, value = 1, step = 0.1),
                  sliderInput("forest_text_size", "Text size:",
                            min = 8, max = 24, value = 16)
                )
              )
            ),
            
            hr(),
            
            fluidRow(
              column(3,
                div(class = "control-section",
                  h4("Plot Annotations", style = "margin-top: 0;"),
                  checkboxInput("show_effect_values", "Show effect sizes", value = TRUE),
                  checkboxInput("show_ci_values", "Show CIs", value = TRUE),
                  checkboxInput("show_weights", "Show weights", value = TRUE),
                  checkboxInput("show_overall_diamond", "Show overall diamond", value = TRUE),
                  checkboxInput("show_overall_pvalue", "Show overall p-value", value = TRUE),
                  checkboxInput("show_group_diamonds", "Show group diamonds", value = TRUE),
                  
                  hr(style = "margin: 8px 0;"),
                  
                  checkboxInput("show_study_details", "Show study details", value = FALSE),
                  conditionalPanel(
                    condition = "input.show_study_details",
                    selectInput("detail_vars", "Detail variables:",
                               choices = c(), multiple = TRUE),
                    helpText("Add variables to show in study labels (e.g., Muscle_Group, Measure)")
                  )
                )
              ),
              
              column(3,
                div(class = "control-section",
                  h4("Line Appearance", style = "margin-top: 0;"),
                  sliderInput("ci_line_thickness", "CI line thickness:",
                            min = 0.3, max = 2, value = 0.5, step = 0.1),
                  sliderInput("point_outline_thickness", "Point outline:",
                            min = 0.3, max = 2, value = 0.8, step = 0.1),
                  sliderInput("line_darkness", "Line darkness:",
                            min = 0.3, max = 1, value = 1, step = 0.1),
                  helpText("Higher = darker, lower = lighter")
                )
              ),
              
              column(3,
                     div(class = "control-section",
                         h4("Order & Grouping", style = "margin-top: 0;"),
                         selectInput("forest_order_var", "Order studies by:",
                                     choices = c("Effect size" = "yi")),
                         radioButtons("forest_order_dir", "Direction:",
                                      choices = c("Descending" = "desc",
                                                  "Ascending"  = "asc"),
                                      selected = "desc", inline = TRUE),
                         selectInput("forest_group_var", "Colour by variable:",
                                     choices = c("(none)" = "")),
                         checkboxInput("show_group_summaries",
                                       "Show group-specific summaries", value = FALSE)
                     ),
                     div(class = "control-section",
                         h4("Plot Size", style = "margin-top: 0;"),
                         sliderInput("forest_width", "Plot width (inches):",
                                     min = 6, max = 20, value = 20, step = 1),
                         sliderInput("forest_height", "Plot height (inches):",
                                     min = 6, max = 30, value = 12, step = 1)
                     )
              ),
              column(3,
                     br(), br(),
                     actionButton("update_forest", "Update Plot",
                                  class = "btn-success btn-lg", icon = icon("refresh"))
              )
            )
          )
        ),
        
        fluidRow(
          box(
            title = "Enhanced Forest Plot",
            status = "primary",
            width = 12,
            
            plotOutput("enhanced_forest", height = "auto"),
            
            conditionalPanel(
              condition = "input.show_summary_stats",
              div(class = "summary-stats-box",
                h4(icon("chart-bar"), " Meta-Analysis Summary Statistics"),
                fluidRow(
                  column(6,
                    verbatimTextOutput("forest_summary_stats")
                  ),
                  column(6,
                    verbatimTextOutput("forest_heterogeneity_stats")
                  )
                )
              )
            ),
            
            # NEW: Sensitivity Analysis Results (shown when enabled)
            conditionalPanel(
              condition = "input.run_rho_sensitivity",
              div(class = "sensitivity-box",
                h4(icon("chart-line"), " Sensitivity Analysis: Effect of Assumed ρ"),
                verbatimTextOutput("sensitivity_output")
              )
            ),
            
            hr(),
            
            downloadButton("download_enhanced_forest", "Download Plot",
                         class = "btn-primary")
          )
        )
      ),
      
      # =================== FUNNEL PLOT TAB =====================================
      tabItem(
        tabName = "funnel",
        
        fluidRow(
          box(
            title = "Funnel Plot",
            status = "warning",
            width = 8,
            
            plotOutput("funnel_plot", height = "500px"),
            
            # Trim-and-fill funnel plot (shown conditionally)
            conditionalPanel(
              condition = "input.run_trimfill && output.egger_significant",
              hr(),
              h4(icon("balance-scale"), " Trim-and-Fill Adjusted Funnel Plot"),
              p("Filled circles represent imputed studies to restore funnel symmetry."),
              plotOutput("trimfill_funnel_plot", height = "500px")
            )
          ),
          
          box(
            title = "Options",
            status = "info",
            width = 4,
            
            checkboxInput("contour_enhanced", "Enhanced contours", value = TRUE),
            checkboxInput("label_funnel", "Label studies", value = TRUE),
            
            conditionalPanel(
              condition = "input.label_funnel",
              sliderInput("label_size", "Label size:",
                        min = 0.3, max = 1.2, value = 0.7, step = 0.1),
              numericInput("max_funnel_labels", "Max labels:",
                         value = 50, min = 5, max = 200)
            ),
            
            hr(),
            
            checkboxInput("run_egger", "Run Egger's test", value = FALSE),
            
            conditionalPanel(
              condition = "input.run_egger",
              verbatimTextOutput("egger_results"),
              
              hr(),
              
              # Trim-and-fill option (only shown after Egger's test)
              checkboxInput("run_trimfill", 
                          "Run Duval & Tweedie's trim-and-fill if Egger's test is significant", 
                          value = TRUE),
              
              conditionalPanel(
                condition = "input.run_trimfill",
                selectInput("trimfill_side", "Trim-and-fill side:",
                          choices = c("Left (suppress positive)" = "left",
                                     "Right (suppress negative)" = "right"),
                          selected = "left"),
                helpText("'Left' assumes small studies with small/negative effects are missing. 
                         'Right' assumes small studies with large/positive effects are missing.")
              )
            ),
            
            # Trim-and-fill results box
            conditionalPanel(
              condition = "input.run_egger && input.run_trimfill",
              hr(),
              div(id = "trimfill_results_box",
                uiOutput("trimfill_results_ui")
              )
            ),
            
            hr(),
            
            actionButton("update_funnel", "Update Plot",
                       class = "btn-warning", icon = icon("refresh"))
          )
        )
      ),
      
      # =================== EXPORT TAB ==========================================
      tabItem(
        tabName = "export",
        
        fluidRow(
          box(
            title = "Methods & Results Write-Up",
            status = "success",
            solidHeader = TRUE,
            width = 12,
            collapsible = TRUE,
            
            p("Generate a publication-ready methods and results section based on your analysis."),
            
            fluidRow(
              column(4,
                checkboxInput("include_methods", "Include Methods section", value = TRUE),
                checkboxInput("include_results", "Include Results section", value = TRUE),
                checkboxInput("include_references", "Include statistical references", value = TRUE)
              ),
              column(4,
                selectInput("writeup_format", "Output format:",
                          choices = c("Plain text" = "txt",
                                     "Markdown" = "md",
                                     "Word document" = "docx")),
                selectInput("citation_style", "Citation style:",
                          choices = c("APA 7th" = "apa",
                                     "Vancouver" = "vancouver"))
              ),
              column(4,
                br(),
                actionButton("generate_writeup", "Generate Write-Up", 
                           class = "btn-success btn-lg", icon = icon("file-alt")),
                br(), br(),
                downloadButton("download_writeup", "Download Write-Up", class = "btn-primary")
              )
            ),
            
            hr(),
            
            h4("Preview:"),
            div(style = "max-height: 500px; overflow-y: auto; background: #f9f9f9; padding: 15px; border-radius: 5px; border: 1px solid #ddd;",
              uiOutput("writeup_preview")
            )
          )
        ),
        
        fluidRow(
          box(
            title = "Export Data",
            status = "info",
            width = 6,
            
            downloadButton("download_filtered", "Download Filtered Data"),
            br(), br(),
            downloadButton("download_es", "Download Effect Sizes"),
            br(), br(),
            downloadButton("download_results", "Download Results Summary")
          ),
          
          box(
            title = "Export Plots",
            status = "primary",
            width = 6,
            
            selectInput("plot_format", "Format:",
                       choices = c("PDF" = "pdf", "PNG" = "png")),
            
            conditionalPanel(
              condition = "input.plot_format == 'png'",
              numericInput("export_dpi", "DPI:", value = 300, min = 72, max = 600)
            ),
            
            downloadButton("download_forest", "Download Forest Plot"),
            br(), br(),
            downloadButton("download_funnel", "Download Funnel Plot"),
            br(), br(),
            downloadButton("download_trimfill_funnel", "Download Trim-and-Fill Funnel Plot")
          )
        )
      )
    )
  )
)

# ========================= SERVER ============================================
server <- function(input, output, session) {
  
  values <- reactiveValues(
    raw_data = NULL,
    filtered_data = NULL,
    es_data = NULL,
    ma_model = NULL,
    ma_model_robust = NULL,        # NEW: Store robust model results
    sensitivity_results = NULL,    # NEW: Store sensitivity analysis results
    V_matrix = NULL,               # NEW: Store vcalc V matrix
    active_filter_vars = character(0),
    available_cat_vars = character(0),
    missing_data_report = NULL,
    forest_model = NULL,
    forest_group_models = NULL,
    trimfill_model = NULL,
    egger_pval = NULL,
    writeup_text = NULL,
    filter_history = list(),
    model_type = NULL              # NEW: Track model type
  )
  
  # =================== DATA LOADING ==========================================
  
  output$sheet_selector <- renderUI({
    req(input$file)
    if(grepl("\\.xlsx$|\\.xls$", input$file$name, ignore.case = TRUE)) {
      sheets <- excel_sheets(input$file$datapath)
      selectInput("sheet", "Select sheet:", choices = sheets)
    }
  })
  
  observeEvent(input$load_data, {
    req(input$file)
    
    tryCatch({
      if(grepl("\\.csv$", input$file$name, ignore.case = TRUE)) {
        values$raw_data <- read.csv(input$file$datapath, stringsAsFactors = FALSE)
      } else {
        req(input$sheet)
        values$raw_data <- read_excel(input$file$datapath, sheet = input$sheet)
      }
      
      values$filtered_data <- values$raw_data
      values$active_filter_vars <- character(0)
      
      values$available_cat_vars <- names(values$raw_data)[sapply(values$raw_data, function(col) {
        is.character(col) || is.factor(col) || length(unique(na.omit(col))) <= 20
      })]
      
      cols <- names(values$raw_data)
      
      # Auto-detect function - finds best match from patterns
      auto_detect <- function(patterns, col_names, must_exist = FALSE) {
        for(pattern in patterns) {
          matches <- grep(pattern, col_names, ignore.case = TRUE, value = TRUE)
          if(length(matches) > 0) return(matches[1])
        }
        if(must_exist) return(col_names[1]) else return("")
      }
      
      # Define search patterns (in priority order)
      studyid_patterns <- c("^studyid$", "^study_id$", "study.?id", "^author$", 
                           "first.?author", "citation", "^id$")
      mean1_patterns <- c("post.?mean", "mean.?post", "mean.?1", "m1i?$", 
                         "follow.?up.?mean", "end.?mean", "final.?mean")
      mean2_patterns <- c("pre.?mean", "mean.?pre", "mean.?2", "m2i?$", 
                         "baseline.?mean", "initial.?mean", "start.?mean")
      sd1_patterns <- c("post.?std", "post.?sd", "std.?post", "sd.?post", 
                       "sd.?1", "s1", "follow.?up.?sd", "end.?sd")
      sd2_patterns <- c("pre.?std", "pre.?sd", "std.?pre", "sd.?pre", 
                       "sd.?2", "s2", "baseline.?sd", "initial.?sd")
      n_patterns <- c("^n$", "^sample.?size$", "sample.?n", "n.?total", 
                     "n1i?$", "participants", "subjects")
      group_patterns <- c("^group$", "^arm$", "condition", "intervention", 
                         "treatment", "^tx$")
      corr_patterns <- c("^r$", "^corr", "correlation", "pre.?post.?r", 
                        "r.?pre.?post", "^r_")
      change_sd_patterns <- c("change.?sd", "sd.?change", "std.?change", 
                             "change.?std", "sd.?diff", "diff.?sd")
      
      # Auto-detect and update selectors
      updateSelectInput(session, "studyid_col", choices = cols,
                       selected = auto_detect(studyid_patterns, cols, must_exist = TRUE))
      updateSelectInput(session, "mean1_col", choices = cols,
                       selected = auto_detect(mean1_patterns, cols, must_exist = TRUE))
      updateSelectInput(session, "mean2_col", choices = cols,
                       selected = auto_detect(mean2_patterns, cols, must_exist = TRUE))
      updateSelectInput(session, "sd1_col", choices = cols,
                       selected = auto_detect(sd1_patterns, cols, must_exist = TRUE))
      updateSelectInput(session, "sd2_col", choices = cols,
                       selected = auto_detect(sd2_patterns, cols, must_exist = TRUE))
      updateSelectInput(session, "n_col", choices = cols,
                       selected = auto_detect(n_patterns, cols, must_exist = TRUE))
      updateSelectInput(session, "corr_col", choices = c("", cols),
                       selected = auto_detect(corr_patterns, cols))
      updateSelectInput(session, "change_sd_col", choices = c("", cols),
                       selected = auto_detect(change_sd_patterns, cols))
      updateSelectInput(session, "group_col", choices = c("(none)" = "", cols),
                       selected = auto_detect(group_patterns, cols))
      updateSelectInput(session, "cat_var_to_add", choices = values$available_cat_vars)
      
      # Update detail variables selector (for forest plot labels)
      detail_candidates <- setdiff(cols, c(input$studyid_col, input$mean1_col, input$mean2_col,
                                          input$sd1_col, input$sd2_col, input$n_col))
      updateSelectInput(session, "detail_vars", choices = detail_candidates)
      
      # Show notification with detection results
      detected_cols <- c(
        paste("Study ID:", auto_detect(studyid_patterns, cols, must_exist = TRUE)),
        paste("Mean 1:", auto_detect(mean1_patterns, cols, must_exist = TRUE)),
        paste("Mean 2:", auto_detect(mean2_patterns, cols, must_exist = TRUE)),
        paste("Group:", ifelse(auto_detect(group_patterns, cols) == "", "Not detected", 
                              auto_detect(group_patterns, cols)))
      )
      
      showNotification(
        HTML(paste("Data loaded with auto-detected columns:<br>", 
                  paste(detected_cols, collapse = "<br>"),
                  "<br><br>Please verify column mappings before calculating.")),
        type = "message",
        duration = 10
      )
      
      showNotification("Data loaded successfully!", type = "message")
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  output$import_summary <- renderPrint({
    req(values$raw_data)
    cat("Rows:", nrow(values$raw_data), "\n")
    cat("Columns:", ncol(values$raw_data), "\n")
  })
  
  output$preview_table <- DT::renderDataTable({
    req(values$raw_data)
    DT::datatable(head(values$raw_data, 100),
                  options = list(scrollX = TRUE, pageLength = 10))
  })
  
  # =================== FILTERING =============================================
  
  observeEvent(input$add_cat_var, {
    req(input$cat_var_to_add)
    var <- input$cat_var_to_add
    if(!(var %in% values$active_filter_vars)) {
      values$active_filter_vars <- c(values$active_filter_vars, var)
      showNotification(paste("Added filter for:", var), type = "message", duration = 2)
    }
  })
  
  observeEvent(input$remove_all_filters, {
    values$active_filter_vars <- character(0)
    showNotification("All filters removed", type = "warning", duration = 2)
  })
  
  output$active_filters_ui <- renderUI({
    req(values$raw_data)
    
    if(length(values$active_filter_vars) == 0) {
      return(div(style = "padding: 20px; text-align: center;",
        p("No active filters.", style = "color: #999; font-style: italic;"),
        p("Select a variable above and click 'Add Variable' to start filtering.", 
          style = "color: #999; font-size: 12px;")
      ))
    }
    
    filter_uis <- lapply(values$active_filter_vars, function(var) {
      if(!(var %in% names(values$raw_data))) return(NULL)
      
      choices <- sort(unique(na.omit(values$raw_data[[var]])))
      
      div(class = "filter-box",
        # Header with variable name and remove button
        div(style = "display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;",
          h5(style = "margin: 0; color: #333; font-weight: bold;", var),
          actionButton(paste0("remove_", var), "Remove",
                     class = "btn-danger btn-sm", 
                     icon = icon("times"),
                     style = "font-size: 11px; padding: 2px 8px;")
        ),
        
        # Instructions
        div(style = "margin-bottom: 8px; color: #666; font-size: 12px;",
          HTML("<strong>Select values to INCLUDE</strong> (uncheck to exclude):")
        ),
        
        # All/None buttons
        div(style = "margin-bottom: 10px;",
          div(class = "btn-group btn-group-sm", role = "group",
            actionButton(paste0("select_all_", var), 
                       HTML("<i class='fa fa-check-square'></i> All"), 
                       class = "btn-default btn-sm",
                       style = "font-size: 11px;"),
            actionButton(paste0("select_none_", var), 
                       HTML("<i class='fa fa-square-o'></i> None"), 
                       class = "btn-default btn-sm",
                       style = "font-size: 11px;")
          ),
          span(style = "margin-left: 10px; color: #999; font-size: 11px;",
            paste0("(", length(choices), " values)"))
        ),
        
        # Checkboxes for values
        div(style = "background: white; padding: 10px; border-radius: 3px; border: 1px solid #ddd;",
          checkboxGroupInput(paste0("filter_cat_", var), 
                            label = NULL,
                            choices = choices, 
                            selected = choices, 
                            inline = TRUE)
        )
      )
    })
    
    do.call(tagList, filter_uis)
  })
  
  observe({
    req(values$active_filter_vars)
    for(var in values$active_filter_vars) {
      local({
        var_local <- var
        observeEvent(input[[paste0("remove_", var_local)]], {
          values$active_filter_vars <- setdiff(values$active_filter_vars, var_local)
        })
        observeEvent(input[[paste0("select_all_", var_local)]], {
          if(var_local %in% names(values$raw_data)) {
            all_choices <- sort(unique(na.omit(values$raw_data[[var_local]])))
            updateCheckboxGroupInput(session, paste0("filter_cat_", var_local),
                                    selected = all_choices)
          }
        })
        observeEvent(input[[paste0("select_none_", var_local)]], {
          updateCheckboxGroupInput(session, paste0("filter_cat_", var_local),
                                  selected = character(0))
        })
      })
    }
  })
  
  observeEvent(input$apply_filters, {
    req(values$raw_data)
    data <- values$raw_data
    filter_details <- list()
    
    for(var in values$active_filter_vars) {
      if(var %in% names(data)) {
        selected <- input[[paste0("filter_cat_", var)]]
        if(!is.null(selected) && length(selected) > 0) {
          data <- data %>% filter(.data[[var]] %in% selected)
          filter_details[[var]] <- selected
        }
      }
    }
    values$filtered_data <- data
    values$filter_history <- filter_details
    showNotification(paste("Filters applied.", nrow(data), "rows remaining."), type = "message")
  })
  
  observeEvent(input$reset_filters, {
    values$filtered_data <- values$raw_data
    values$filter_history <- list()
    showNotification("Data reset to original.", type = "warning")
  })
  
  output$filter_summary <- renderPrint({
    req(values$raw_data, values$filtered_data)
    cat("Original rows:", nrow(values$raw_data), "\n")
    cat("Filtered rows:", nrow(values$filtered_data), "\n")
    cat("Excluded rows:", nrow(values$raw_data) - nrow(values$filtered_data), "\n")
  })
  
  output$active_vars_list <- renderPrint({
    if(length(values$active_filter_vars) == 0) {
      cat("None\n")
    } else {
      cat(paste0(seq_along(values$active_filter_vars), ". ", 
                values$active_filter_vars, "\n"), sep = "")
    }
  })
  
  output$n_total <- renderValueBox({
    req(values$raw_data)
    valueBox(nrow(values$raw_data), "Total Rows", icon = icon("database"), color = "blue")
  })
  
  output$n_filtered <- renderValueBox({
    req(values$filtered_data)
    valueBox(nrow(values$filtered_data), "After Filtering", icon = icon("filter"), color = "green")
  })
  
  output$n_excluded <- renderValueBox({
    req(values$raw_data, values$filtered_data)
    excluded <- nrow(values$raw_data) - nrow(values$filtered_data)
    valueBox(excluded, "Excluded", icon = icon("ban"), color = if(excluded > 0) "red" else "light-blue")
  })
  
  output$filtered_preview <- DT::renderDataTable({
    req(values$filtered_data)
    DT::datatable(values$filtered_data, options = list(scrollX = TRUE, pageLength = 25))
  })
  
  # =================== EFFECT SIZE CALCULATION (condensed) ===================
  
  output$precalc_check <- renderPrint({
    req(values$filtered_data)
    data <- values$filtered_data
    cat("Data check:\n- Rows:", nrow(data), "\n\n")
    
    check_col <- function(col_input, col_name) {
      if(!is.null(col_input) && col_input != "" && col_input %in% names(data)) {
        n_missing <- sum(is.na(data[[col_input]]))
        if(n_missing > 0) {
          cat("- ", col_name, ": ", col_input, " ⚠ (", n_missing, " missing)\n", sep = "")
        } else {
          cat("- ", col_name, ": ", col_input, " ✓\n", sep = "")
        }
        return(n_missing)
      }
      return(0)
    }
    
    total_missing <- check_col(input$studyid_col, "Study ID")
    total_missing <- total_missing + check_col(input$mean1_col, "Mean 1")
    total_missing <- total_missing + check_col(input$mean2_col, "Mean 2")
    total_missing <- total_missing + check_col(input$sd1_col, "SD 1")
    total_missing <- total_missing + check_col(input$n_col, "Sample size")
    
    if(input$corr_source == "compute") {
      total_missing <- total_missing + check_col(input$change_sd_col, "Change SD")
      total_missing <- total_missing + check_col(input$sd2_col, "SD 2")
    }
    
    cat("\n")
    if(total_missing > 0) {
      cat("⚠ WARNING:", total_missing, "missing values\n")
      if(input$auto_remove_missing) cat("→ Will auto-remove incomplete rows\n")
    } else {
      cat("✓ Ready to calculate!\n")
    }
    
    values$missing_data_report <- list(total_missing = total_missing, n_rows = nrow(data))
  })
  
  output$missing_data_warning <- renderUI({
    req(values$missing_data_report)
    if(values$missing_data_report$total_missing > 0) {
      div(class = "missing-data-warning",
        h4(icon("exclamation-triangle"), " Missing Data Detected"),
        p(strong(values$missing_data_report$total_missing), " missing values found."),
        p(if(input$auto_remove_missing) "Incomplete rows will be removed." 
          else HTML("<span style='color: red;'>Calculation will fail.</span>"))
      )
    }
  })
  
  observeEvent(input$calculate_es, {
    req(values$filtered_data, input$studyid_col, input$mean1_col, input$sd1_col, input$n_col)
    
    tryCatch({
      data <- values$filtered_data
      
      studyid <- as.character(data[[input$studyid_col]])
      mean1 <- as.numeric(data[[input$mean1_col]])
      sd1 <- as.numeric(data[[input$sd1_col]])
      n <- as.numeric(data[[input$n_col]])
      
      if(!is.null(input$group_col) && input$group_col != "" && input$group_col %in% names(data)) {
        group <- as.character(data[[input$group_col]])
      } else {
        group <- rep("All", nrow(data))
      }
      
      required_complete <- complete.cases(studyid, mean1, sd1, n)
      
      if(input$es_type %in% c("SMCR", "MD")) {
        req(input$mean2_col)
        mean2 <- as.numeric(data[[input$mean2_col]])
        required_complete <- required_complete & !is.na(mean2)
      }
      
      if(input$es_type == "SMD") {
        req(input$mean2_col, input$sd2_col)
        mean2 <- as.numeric(data[[input$mean2_col]])
        sd2 <- as.numeric(data[[input$sd2_col]])
        required_complete <- required_complete & !is.na(mean2) & !is.na(sd2)
      }
      
      n_incomplete <- sum(!required_complete)
      
      if(n_incomplete > 0) {
        if(input$auto_remove_missing) {
          showNotification(paste("Removing", n_incomplete, "rows with missing data..."), 
                          type = "warning", duration = 5)
          data <- data[required_complete, ]
          studyid <- studyid[required_complete]
          mean1 <- mean1[required_complete]
          sd1 <- sd1[required_complete]
          n <- n[required_complete]
          group <- group[required_complete]
          if(exists("mean2")) mean2 <- mean2[required_complete]
          if(exists("sd2")) sd2 <- sd2[required_complete]
        } else {
          stop(paste(n_incomplete, "rows have missing data."))
        }
      }
      
      # Correlation
      if(input$corr_source == "fixed") {
        r_prepost <- rep(input$corr_value, nrow(data))
      } else if(input$corr_source == "column" && !is.null(input$corr_col) && input$corr_col != "") {
        r_prepost <- as.numeric(data[[input$corr_col]])
        r_prepost[is.na(r_prepost) | abs(r_prepost) > 1] <- 0.5
      } else if(input$corr_source == "compute") {
        req(input$change_sd_col, input$sd2_col)
        sd2 <- as.numeric(data[[input$sd2_col]])
        sd_change <- as.numeric(data[[input$change_sd_col]])
        r_prepost <- (sd1^2 + sd2^2 - sd_change^2) / (2 * sd1 * sd2)
        invalid <- which(is.na(r_prepost) | abs(r_prepost) > 1)
        if(length(invalid) > 0) r_prepost[invalid] <- 0.5
        r_prepost <- pmax(-0.99, pmin(0.99, r_prepost))
      } else {
        r_prepost <- rep(0.5, nrow(data))
      }
      
      # Calculate
      if(input$es_type == "SMCR") {
        es_result <- escalc(measure = "SMCR", m1i = mean1, m2i = mean2,
                          sd1i = sd1, ni = n, ri = r_prepost, data = data)
      } else if(input$es_type == "SMD") {
        es_result <- escalc(measure = "SMD", m1i = mean1, m2i = mean2,
                          sd1i = sd1, sd2i = sd2, n1i = n, n2i = n, data = data)
      } else if(input$es_type == "MD") {
        es_result <- escalc(measure = "MD", m1i = mean1, m2i = mean2,
                          sd1i = sd1, ni = n, data = data)
      }
      
      if(input$apply_hedges) {
        df <- n - 1
        J <- 1 - 3 / (4 * df - 1)
        es_result$yi <- es_result$yi * J
        es_result$vi <- es_result$vi * J^2
      }
      
      es_result$StudyID <- studyid
      es_result$Group <- group
      es_result$n <- n
      es_result$se <- sqrt(es_result$vi)
      es_result$ci_lower <- es_result$yi - 1.96 * es_result$se
      es_result$ci_upper <- es_result$yi + 1.96 * es_result$se
      
      if(input$corr_source == "compute") es_result$r_computed <- r_prepost
      
      
      # ADD UNIQUE EFFECT SIZE IDs FOR THREE-LEVEL MODEL
      es_result$EffectSizeID <- 1:nrow(es_result)
      
      # Create a combined study-group ID for proper nesting
      es_result$StudyGroup <- paste(es_result$StudyID, es_result$Group, sep = "_")
      
      values$es_data <- as.data.frame(es_result)
      
      showNotification(paste("✓ Calculated", nrow(es_result), "effect sizes!"), 
                      type = "message", duration = 5)
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error", duration = 15)
    })
  })
  
  output$es_table <- DT::renderDataTable({
    req(values$es_data)
    display_cols <- c("StudyID", "Group", "yi", "vi", "se", "ci_lower", "ci_upper", "n")
    if("r_computed" %in% names(values$es_data)) display_cols <- c(display_cols, "r_computed")
    display_data <- values$es_data[, intersect(display_cols, names(values$es_data))]
    DT::datatable(display_data, options = list(scrollX = TRUE, pageLength = 25)) %>%
      DT::formatRound(setdiff(names(display_data), c("StudyID", "Group", "n")), 3)
  })
  
  output$es_distribution <- renderPlot({
    req(values$es_data)
    ggplot(values$es_data, aes(x = yi)) +
      geom_histogram(bins = 30, fill = "#3498db", alpha = 0.7, color = "white") +
      geom_vline(xintercept = 0, linetype = "dashed", color = "red") +
      geom_vline(xintercept = mean(values$es_data$yi), color = "darkblue", size = 1) +
      labs(x = "Effect Size", y = "Frequency") + theme_minimal()
  })
  
  output$es_summary <- renderPrint({
    req(values$es_data)
    cat("n =", nrow(values$es_data), "\n")
    cat("Mean =", round(mean(values$es_data$yi), 3), "\n")
    cat("SD =", round(sd(values$es_data$yi), 3), "\n")
    cat("Range: [", round(min(values$es_data$yi), 3), ",", 
        round(max(values$es_data$yi), 3), "]\n")
  })
  
  # Update forest order & grouping controls when effect sizes are available
  observeEvent(values$es_data, {
    req(values$es_data)
    cols_es <- names(values$es_data)
    
    # Get the effect size type label
    es_type_label <- if(!is.null(input$es_type) && input$es_type != "") {
      paste0("Effect size (", input$es_type, ")")
    } else {
      "Effect size"
    }
    
    # Build order choices with proper named vector
    # First entry is the effect size with dynamic label
    order_choices <- c("yi", "StudyID", setdiff(cols_es, c("yi", "StudyID")))
    names(order_choices) <- c(es_type_label, "StudyID", setdiff(cols_es, c("yi", "StudyID")))
    
    updateSelectInput(session, "forest_order_var",
                      choices  = order_choices,
                      selected = "yi")
    
    # Grouping variable: any column from es_data
    # Also add TreatmentSubgroup as an option (will only work when split treatment is used)
    group_choices <- c("(none)" = "", "TreatmentSubgroup (for split treatment)" = "TreatmentSubgroup", cols_es)
    updateSelectInput(session, "forest_group_var",
                      choices  = group_choices,
                      selected = "")
    
    # Update between-group difference controls
    if("Group" %in% names(values$es_data)) {
      groups_available <- unique(as.character(values$es_data$Group))
      
      # Try to intelligently select defaults
      # Look for training-like group
      training_default <- groups_available[grepl("train|treat|exercise", groups_available, ignore.case = TRUE)][1]
      if(is.na(training_default)) training_default <- groups_available[1]
      
      # Look for control-like group  
      control_default <- groups_available[grepl("control|immobil", groups_available, ignore.case = TRUE)][1]
      if(is.na(control_default)) control_default <- groups_available[min(2, length(groups_available))]
      
      updateSelectInput(session, "diff_group1",
                        choices = groups_available,
                        selected = training_default)
      updateSelectInput(session, "diff_group2", 
                        choices = groups_available,
                        selected = control_default)
    }
    
    # Update matching variables for between-group differences
    # Suggest columns that are likely outcome identifiers
    potential_match_cols <- c("Measure", "Muscle_Group", "Measure_Type", "Measure_Method", 
                              "Limb", "Trained_Untrained_Muscle_Group", "Outcome")
    available_match_cols <- intersect(potential_match_cols, cols_es)
    
    # Also include any other character/factor columns that might be useful
    other_cols <- cols_es[sapply(values$es_data[cols_es], function(x) {
      is.character(x) || is.factor(x)
    })]
    other_cols <- setdiff(other_cols, c("StudyID", "Group", "StudyGroup", "label", available_match_cols))
    
    all_match_choices <- c(available_match_cols, other_cols)
    
    # Don't pre-select any match variables - user should explicitly choose
    # This prevents confusion about when outcome labels appear
    updateSelectInput(session, "diff_match_vars",
                      choices = all_match_choices,
                      selected = character(0))  # Empty selection by default
    
    # Update treatment subgroup variable choices
    # Prioritize Training_Type and similar columns
    priority_subgroup_cols <- c("Training_Type", "Training_Intensity", "Muscle_Training_Target")
    available_subgroup_cols <- intersect(priority_subgroup_cols, cols_es)
    other_subgroup_cols <- setdiff(other_cols, available_subgroup_cols)
    subgroup_choices <- c("(none)" = "", available_subgroup_cols, other_subgroup_cols)
    
    # Default to Training_Type if available
    default_subgroup <- if("Training_Type" %in% cols_es) "Training_Type" else ""
    
    updateSelectInput(session, "diff_treatment_subgroup",
                      choices = subgroup_choices,
                      selected = "")  # Don't auto-select, let user choose
  })

  # =================== BASIC META-ANALYSIS ===================================
  
  observeEvent(input$run_analysis, {
    req(values$es_data)
    tryCatch({
      # Ensure we have the required columns
      if(!"EffectSizeID" %in% names(values$es_data)) {
        values$es_data$EffectSizeID <- 1:nrow(values$es_data)
      }
      if(!"StudyGroup" %in% names(values$es_data)) {
        values$es_data$StudyGroup <- paste(values$es_data$StudyID, 
                                            values$es_data$Group, sep = "_")
      }
      
      # DETERMINE IF THREE-LEVEL MODEL IS NEEDED
      # Count effect sizes per study-group
      es_per_studygroup <- table(values$es_data$StudyGroup)
      n_studygroups <- length(es_per_studygroup)
      n_effects <- nrow(values$es_data)
      max_es_per_sg <- max(es_per_studygroup)
      n_sg_with_multiple <- sum(es_per_studygroup > 1)
      
      # Use three-level model only if there are multiple effect sizes within at least some study-groups
      use_three_level <- n_effects > n_studygroups && n_sg_with_multiple > 0
      
      # ========= CHE MODEL WITH vcalc() =========
      if(input$use_che_model && use_three_level) {
        # Create V matrix using vcalc() with assumed rho
        # This accounts for correlated sampling errors within studies
        values$V_matrix <- vcalc(
          vi = values$es_data$vi,
          cluster = values$es_data$StudyGroup,
          obs = values$es_data$EffectSizeID,
          data = values$es_data,
          rho = input$assumed_rho
        )
        
        # Fit model with V matrix (Correlated and Hierarchical Effects model)
        values$ma_model <- rma.mv(
          yi = yi, 
          V = values$V_matrix,  # Use V matrix instead of vi vector
          random = ~ 1 | StudyGroup/EffectSizeID,
          data = values$es_data,
          method = input$ma_model,
          test = ifelse(input$test_type == "knha", "t", "z")
        )
        
        values$model_type <- "three_level_che"
        
        showNotification(
          paste0("CHE model fitted with ρ = ", input$assumed_rho, 
                 " (", n_sg_with_multiple, " study-groups have multiple effect sizes)"), 
          type = "message", duration = 5
        )
        
      } else if(use_three_level) {
        # Standard three-level model (assumes independent sampling errors)
        values$V_matrix <- NULL
        
        values$ma_model <- rma.mv(
          yi = yi, 
          V = vi,
          random = ~ 1 | StudyGroup/EffectSizeID,
          data = values$es_data,
          method = input$ma_model,
          test = ifelse(input$test_type == "knha", "t", "z")
        )
        
        # Store model type for downstream functions
        values$model_type <- "three_level"
        
        showNotification(
          paste0("Three-level meta-analysis complete! (", n_sg_with_multiple, 
                 " of ", n_studygroups, " study-groups have multiple effect sizes)"), 
          type = "message", duration = 5
        )
        
      } else {
        # STANDARD TWO-LEVEL MODEL
        # Only one effect size per study-group, so three-level model is unnecessary
        values$V_matrix <- NULL
        
        test_arg <- ifelse(input$test_type == "knha", "knha", "z")
        
        values$ma_model <- rma(
          yi = yi, 
          vi = vi,
          data = values$es_data,
          method = input$ma_model,
          test = test_arg
        )
        
        # Store model type for downstream functions
        values$model_type <- "two_level"
        
        showNotification(
          paste0("Two-level meta-analysis complete! (Each of ", n_studygroups, 
                 " study-groups contributes one effect size)"), 
          type = "message", duration = 5
        )
      }
      
      # ========= ROBUST VARIANCE ESTIMATION =========
      if(input$use_robust_ve && inherits(values$ma_model, "rma.mv")) {
        # Apply cluster-robust SEs with clubSandwich small-sample corrections
        values$ma_model_robust <- tryCatch({
          robust(values$ma_model, 
                 cluster = values$es_data$StudyGroup, 
                 clubSandwich = TRUE)
        }, error = function(e) {
          showNotification(paste("Robust VE failed:", e$message), type = "warning")
          NULL
        })
        
        if(!is.null(values$ma_model_robust)) {
          showNotification("Cluster-robust SEs computed (clubSandwich)", type = "message", duration = 3)
        }
      } else {
        values$ma_model_robust <- NULL
      }
      
      # ========= SENSITIVITY ANALYSIS =========
      # Run when: sensitivity checkbox is checked AND (data needs three-level OR CHE model is enabled)
      run_sensitivity <- input$run_rho_sensitivity && (use_three_level || input$use_che_model)
      
      if(run_sensitivity) {
        rho_values <- c(0.0, 0.3, 0.5, 0.7, 0.9)
        
        withProgress(message = 'Running sensitivity analysis...', value = 0, {
          sens_results <- lapply(seq_along(rho_values), function(i) {
            rho <- rho_values[i]
            incProgress(1/length(rho_values), detail = paste("ρ =", rho))
            
            V_sens <- tryCatch({
              vcalc(
                vi = values$es_data$vi,
                cluster = values$es_data$StudyGroup,
                obs = values$es_data$EffectSizeID,
                data = values$es_data,
                rho = rho
              )
            }, error = function(e) NULL)
            
            if(is.null(V_sens)) return(NULL)
            
            model_sens <- tryCatch({
              rma.mv(
                yi = yi, 
                V = V_sens,
                random = ~ 1 | StudyGroup/EffectSizeID,
                data = values$es_data,
                method = input$ma_model,
                test = ifelse(input$test_type == "knha", "t", "z")
              )
            }, error = function(e) NULL)
            
            if(is.null(model_sens)) return(NULL)
            
            # Get robust estimates if possible
            robust_sens <- tryCatch({
              robust(model_sens, 
                     cluster = values$es_data$StudyGroup, 
                     clubSandwich = TRUE)
            }, error = function(e) NULL)
            
            if(!is.null(robust_sens)) {
              data.frame(
                rho = rho,
                estimate = round(coef(model_sens), 3),
                se_model = round(model_sens$se, 3),
                se_robust = round(as.numeric(robust_sens$se), 3),
                ci_lb = round(as.numeric(robust_sens$ci.lb), 3),
                ci_ub = round(as.numeric(robust_sens$ci.ub), 3),
                p_robust = as.numeric(robust_sens$pval)
              )
            } else {
              data.frame(
                rho = rho,
                estimate = round(coef(model_sens), 3),
                se_model = round(model_sens$se, 3),
                se_robust = NA,
                ci_lb = round(model_sens$ci.lb, 3),
                ci_ub = round(model_sens$ci.ub, 3),
                p_robust = model_sens$pval
              )
            }
          })
        })
        
        values$sensitivity_results <- do.call(rbind, sens_results[!sapply(sens_results, is.null)])
        
        if(!is.null(values$sensitivity_results) && nrow(values$sensitivity_results) > 0) {
          showNotification("Sensitivity analysis complete", type = "message", duration = 3)
        }
        
      } else {
        values$sensitivity_results <- NULL
      }
      
    }, error = function(e) {
      showNotification(paste("Error:", e$message), type = "error", duration = 10)
    })
  })
  
  output$ma_summary <- renderPrint({
    req(values$ma_model)
    
    if(inherits(values$ma_model, "rma.mv")) {
      # Three-level model output
      if(values$model_type == "three_level_che") {
        cat("=== THREE-LEVEL META-ANALYSIS (CHE Model) ===\n")
        cat("Correlated sampling errors assumed with ρ =", input$assumed_rho, "\n\n")
      } else {
        cat("=== THREE-LEVEL META-ANALYSIS RESULTS ===\n")
        cat("(Independent sampling errors assumed)\n\n")
      }
      
      cat("Overall Effect Estimate:\n")
      cat("  Estimate:", round(coef(values$ma_model), 3), "\n")
      cat("  SE:", round(values$ma_model$se, 3), "\n")
      cat("  95% CI: [", round(values$ma_model$ci.lb, 3), ",", 
          round(values$ma_model$ci.ub, 3), "]\n")
      
      # Test statistic (z or t depending on test type)
      if(!is.null(values$ma_model$zval)) {
        cat("  z =", round(values$ma_model$zval, 3), "\n")
      } else if(!is.null(values$ma_model$QM)) {
        cat("  t =", round(sqrt(values$ma_model$QM), 3), "\n")
      }
      
      cat("  p =", format.pval(values$ma_model$pval, digits = 3), "\n")
      
      # Show robust results inline if available
      if(!is.null(values$ma_model_robust)) {
        cat("\n--- Cluster-Robust Results (clubSandwich) ---\n")
        cat("  Robust SE:", round(as.numeric(values$ma_model_robust$se), 3), "\n")
        cat("  95% CI: [", round(as.numeric(values$ma_model_robust$ci.lb), 3), ",", 
            round(as.numeric(values$ma_model_robust$ci.ub), 3), "]\n")
        if(!is.null(values$ma_model_robust$dfs)) {
          cat("  df:", round(as.numeric(values$ma_model_robust$dfs), 1), "\n")
        }
        cat("  p =", format.pval(as.numeric(values$ma_model_robust$pval), digits = 3), "\n")
      }
      
    } else {
      # Standard two-level model output
      cat("=== TWO-LEVEL META-ANALYSIS RESULTS ===\n")
      cat("(One effect size per study; three-level model not needed)\n\n")
      cat("Overall Effect Estimate:\n")
      cat("  Estimate:", round(coef(values$ma_model), 3), "\n")
      cat("  SE:", round(values$ma_model$se, 3), "\n")
      cat("  95% CI: [", round(values$ma_model$ci.lb, 3), ",", 
          round(values$ma_model$ci.ub, 3), "]\n")
      
      # Test statistic
      if(!is.null(values$ma_model$zval)) {
        if(input$test_type == "knha") {
          cat("  t(", values$ma_model$dfs, ") =", round(values$ma_model$zval, 3), "\n")
        } else {
          cat("  z =", round(values$ma_model$zval, 3), "\n")
        }
      }
      cat("  p =", format.pval(values$ma_model$pval, digits = 3), "\n")
    }
  })
  
  
  output$heterogeneity <- renderPrint({
    req(values$ma_model)
    
    if(inherits(values$ma_model, "rma.mv")) {
      # Three-level model variance components
      cat("=== VARIANCE COMPONENTS ===\n\n")
      
      # Extract variance components
      sigma2_level3 <- values$ma_model$sigma2[1]  # Between study-groups
      sigma2_level2 <- values$ma_model$sigma2[2]  # Within study-groups
      
      cat("Level 3 (Between Study-Groups):\n")
      cat("  σ² =", round(sigma2_level3, 4), "\n\n")
      
      cat("Level 2 (Within Study-Groups):\n")
      cat("  σ² =", round(sigma2_level2, 4), "\n\n")
      
      # Total heterogeneity
      total_var <- sigma2_level3 + sigma2_level2
      cat("Total τ² =", round(total_var, 4), "\n\n")
      
      # Calculate I² for three-level model
      typical_vi <- mean(values$es_data$vi, na.rm = TRUE)
      I2_total <- 100 * total_var / (total_var + typical_vi)
      I2_level3 <- 100 * sigma2_level3 / (total_var + typical_vi)
      I2_level2 <- 100 * sigma2_level2 / (total_var + typical_vi)
      
      cat("I² (Total heterogeneity) =", round(I2_total, 1), "%\n")
      cat("I² (Level 3) =", round(I2_level3, 1), "%\n")
      cat("I² (Level 2) =", round(I2_level2, 1), "%\n\n")
      
      # Omnibus test
      cat("Omnibus Test:\n")
      cat("  Q =", round(values$ma_model$QE, 2), "\n")
      cat("  df =", values$ma_model$QEdf, "\n")
      cat("  p =", format.pval(values$ma_model$QEp, digits = 3), "\n")
      
    } else {
      # Standard two-level model output
      cat("=== HETEROGENEITY ===\n\n")
      cat("τ² =", round(values$ma_model$tau2, 4), "\n")
      cat("τ (SD) =", round(sqrt(values$ma_model$tau2), 4), "\n\n")
      cat("I² =", round(values$ma_model$I2, 1), "%\n")
      cat("H² =", round(values$ma_model$H2, 2), "\n\n")
      cat("Test for Heterogeneity:\n")
      cat("  Q =", round(values$ma_model$QE, 2), "\n")
      cat("  df =", values$ma_model$k - values$ma_model$p, "\n")
      cat("  p =", format.pval(values$ma_model$QEp, digits = 3), "\n")
    }
  })
  
  # NEW: Sensitivity Analysis Output
  output$sensitivity_output <- renderPrint({
    if(is.null(values$sensitivity_results) || nrow(values$sensitivity_results) == 0) {
      cat("Sensitivity analysis not run.\n\n")
      cat("To run sensitivity analysis:\n")
      cat("  1. Check 'Use CHE working model (vcalc)' checkbox\n")
      cat("  2. Check 'Run sensitivity analysis' checkbox\n")
      cat("  3. Click 'Run Meta-Analysis' button (orange button above)\n")
      cat("\nNote: Sensitivity analysis tests how your results change\n")
      cat("with different assumed correlations (ρ = 0.0 to 0.9).\n")
      return()
    }
    
    cat("=== SENSITIVITY TO ASSUMED ρ (within-study correlation) ===\n\n")
    
    # Format p-values for display
    display_results <- values$sensitivity_results
    display_results$p_robust <- sapply(display_results$p_robust, function(p) {
      if(is.na(p)) return("NA")
      if(p < 0.001) return("<0.001")
      return(sprintf("%.3f", p))
    })
    
    print(display_results, row.names = FALSE)
    
    cat("\n--- Interpretation ---\n")
    
    est_range <- range(values$sensitivity_results$estimate, na.rm = TRUE)
    est_diff <- diff(est_range)
    
    if(est_diff < 0.05) {
      cat("✓ STABLE: Estimates vary by only", round(est_diff, 3), "\n")
      cat("  Conclusions are robust to the correlation assumption.\n")
    } else if(est_diff < 0.15) {
      cat("△ MODERATE: Estimates range from", round(est_range[1], 3), "to", round(est_range[2], 3), "\n")
      cat("  Consider reporting results for multiple ρ values.\n")
    } else {
      cat("⚠ SENSITIVE: Estimates range from", round(est_range[1], 3), "to", round(est_range[2], 3), "\n")
      cat("  Interpretation should acknowledge this uncertainty.\n")
    }
    
    # Check CI stability
    if(!all(is.na(values$sensitivity_results$ci_lb))) {
      all_positive <- all(values$sensitivity_results$ci_lb > 0, na.rm = TRUE)
      all_negative <- all(values$sensitivity_results$ci_ub < 0, na.rm = TRUE)
      
      if(all_positive || all_negative) {
        cat("\n✓ Direction of effect is consistent across all ρ values.\n")
      } else {
        cat("\n⚠ Statistical significance varies across ρ values.\n")
      }
    }
  })
  
  
  output$forest_basic <- renderPlot({
    req(values$ma_model, values$es_data)
    forest(values$ma_model, slab = paste(values$es_data$StudyID, values$es_data$Group, sep = " - "),
          xlab = "Effect Size")
  })
  
  # =================== ENHANCED FOREST PLOT ==================================
  
  prepare_forest_data <- reactive({
    req(values$es_data)
    
    if(input$forest_data_type == "arm") {
      data <- values$es_data
      
      # Build labels with optional detail variables
      detail_vars <- input$detail_vars
      
      if(!is.null(detail_vars) && length(detail_vars) > 0 && input$show_study_details) {
        # User wants to show study details - add selected columns to labels
        available_details <- intersect(detail_vars, names(data))
        
        if(length(available_details) > 0) {
          # Create detail string for each row
          data$label <- sapply(1:nrow(data), function(i) {
            base_label <- paste(data$StudyID[i], data$Group[i], sep = " - ")
            
            detail_parts <- sapply(available_details, function(col) {
              val <- data[i, col]
              if(!is.na(val) && val != "") as.character(val) else NA
            })
            detail_parts <- detail_parts[!is.na(detail_parts)]
            
            if(length(detail_parts) > 0) {
              paste0(base_label, "  |  ", paste(detail_parts, collapse = "  |  "))
            } else {
              base_label
            }
          })
          
          # Store for header construction
          data$match_vars_used <- paste(available_details, collapse = "|")
        } else {
          # No valid detail columns
          data$label <- paste(data$StudyID, data$Group, sep = " - ")
          data$match_vars_used <- ""
        }
      } else {
        # No details requested - simple labels
        data$label <- paste(data$StudyID, data$Group, sep = " - ")
        data$match_vars_used <- ""
      }
      
      test_arg <- ifelse(input$test_type == "knha", "knha", "z")
      if(!"EffectSizeID" %in% names(data)) {
        data$EffectSizeID <- 1:nrow(data)
      }
      if(!"StudyGroup" %in% names(data)) {
        data$StudyGroup <- paste(data$StudyID, data$Group, sep = "_")
      }
      
      # DETERMINE IF THREE-LEVEL MODEL IS NEEDED FOR FOREST PLOT
      es_per_studygroup <- table(data$StudyGroup)
      n_studygroups <- length(es_per_studygroup)
      n_effects <- nrow(data)
      n_sg_with_multiple <- sum(es_per_studygroup > 1)
      
      # Use three-level model only if there are multiple effect sizes within study-groups
      use_three_level <- n_effects > n_studygroups && n_sg_with_multiple > 0
      
      if(use_three_level) {
        # Three-level model
        model <- rma.mv(
          yi = yi, 
          V = vi,
          random = ~ 1 | StudyGroup/EffectSizeID,
          slab = label,  # Store labels for funnel plot
          data = data,
          method = input$ma_model,
          test = ifelse(input$test_type == "knha", "t", "z")
        )
      } else {
        # Two-level model
        model <- rma(
          yi = yi, 
          vi = vi,
          slab = label,  # Store labels for funnel plot
          data = data,
          method = input$ma_model,
          test = test_arg
        )
      }
      
      list(data = data, model = model, 
           plot_title = "Forest Plot: Arm-Level Effects",
           subtitle = "Each row shows one arm's effect size")
      
    } else {
      # Between-group differences
      if(!"Group" %in% names(values$es_data) || length(unique(values$es_data$Group)) < 2) {
        showNotification("Cannot compute between-group differences: No Group variable or <2 groups", 
                        type = "error", duration = 10)
        return(NULL)
      }
      
      # USE USER-SELECTED GROUPS (or fall back to auto-detection)
      if(!is.null(input$diff_group1) && !is.null(input$diff_group2) && 
         input$diff_group1 != "" && input$diff_group2 != "" &&
         input$diff_group1 != input$diff_group2) {
        group1 <- input$diff_group1
        group2 <- input$diff_group2
      } else {
        # Fall back to auto-detection
        groups_available <- unique(values$es_data$Group)
        
        if(length(groups_available) == 2) {
          group1 <- groups_available[1]
          group2 <- groups_available[2]
          if(grepl("control|immobil", group1, ignore.case = TRUE)) {
            temp <- group1
            group1 <- group2
            group2 <- temp
          }
        } else {
          control_group <- groups_available[grepl("control|immobil", groups_available, ignore.case = TRUE)][1]
          training_group <- groups_available[grepl("train|treat|exercise", groups_available, ignore.case = TRUE)][1]
          
          if(is.na(control_group) || is.na(training_group)) {
            showNotification("Please select Treatment and Control groups from the dropdowns.", 
                            type = "error", duration = 10)
            return(NULL)
          }
          
          group1 <- training_group
          group2 <- control_group
        }
      }
      
      # Verify selected groups exist in data
      if(!(group1 %in% values$es_data$Group) || !(group2 %in% values$es_data$Group)) {
        showNotification("Selected groups not found in data.", 
                        type = "error", duration = 10)
        return(NULL)
      }
      
      # CHECK FOR TREATMENT SUBGROUP VARIABLE (e.g., Training_Type)
      treatment_subgroup_var <- input$diff_treatment_subgroup
      use_treatment_subgroup <- !is.null(treatment_subgroup_var) && 
                                 treatment_subgroup_var != "" && 
                                 treatment_subgroup_var %in% names(values$es_data)
      
      # Check aggregation method
      use_multiple <- !is.null(input$diff_aggregation) && input$diff_aggregation == "multiple"
      
      if(use_multiple || use_treatment_subgroup) {
        # ============ MULTIPLE DIFFERENCES PER STUDY ============
        # Either by outcome matching OR by treatment subgroup OR both
        
        es_data <- values$es_data
        
        # Get matching variables for outcomes - ONLY if "Multiple per study" is selected
        if(use_multiple) {
          match_vars <- input$diff_match_vars
          
          # Only use match_vars if user explicitly selected them
          if(is.null(match_vars) || length(match_vars) == 0) {
            # User wants multiple outcomes but didn't specify which columns to match by
            # Show warning and don't create outcome labels
            showNotification(
              "Please select 'Match outcomes by' variables to differentiate multiple outcomes per study.",
              type = "warning",
              duration = 8
            )
            match_vars <- c()
            es_data$OutcomeKey <- "All"
            es_data$OutcomeLabel <- ""
          } else {
            # User selected specific match variables - use them
            es_data$OutcomeKey <- apply(es_data[, match_vars, drop = FALSE], 1, 
                                        function(x) paste(x, collapse = "_"))
            es_data$OutcomeLabel <- apply(es_data[, match_vars, drop = FALSE], 1, 
                                          function(x) paste(x[!is.na(x) & x != ""], collapse = " | "))
          }
          
        } else {
          # "One per study (averaged)" - DON'T create outcome labels since averaging across outcomes
          match_vars <- c()
          es_data$OutcomeKey <- "All"
          es_data$OutcomeLabel <- ""  # Empty - no outcome-specific labels when averaging!
        }
        
        # Separate treatment and control data
        treatment_data <- es_data %>% filter(Group == group1)
        control_data <- es_data %>% filter(Group == group2)
        
        if(nrow(treatment_data) == 0 || nrow(control_data) == 0) {
          showNotification("No data found for one or both groups", type = "error", duration = 10)
          return(NULL)
        }
        
        # If using treatment subgroup, add it to the treatment data identifier
        if(use_treatment_subgroup) {
          treatment_data$TreatmentSubgroup <- as.character(treatment_data[[treatment_subgroup_var]])
          # Control doesn't have treatment subgroup - will match to all
        }
        
        # Build the pairing:
        # For each unique combination of StudyID + OutcomeKey + TreatmentSubgroup (if used)
        # find matching control and calculate difference
        
        results_list <- list()
        
        # Get unique treatment combinations
        if(use_treatment_subgroup) {
          treatment_combos <- treatment_data %>%
            group_by(StudyID, OutcomeKey, TreatmentSubgroup) %>%
            summarise(
              yi = mean(yi, na.rm = TRUE),
              vi = mean(vi, na.rm = TRUE),
              OutcomeLabel = first(OutcomeLabel),
              .groups = "drop"
            )
        } else {
          treatment_combos <- treatment_data %>%
            group_by(StudyID, OutcomeKey) %>%
            summarise(
              yi = mean(yi, na.rm = TRUE),
              vi = mean(vi, na.rm = TRUE),
              OutcomeLabel = first(OutcomeLabel),
              .groups = "drop"
            )
          treatment_combos$TreatmentSubgroup <- NA
        }
        
        # Get control averages by StudyID + OutcomeKey
        control_combos <- control_data %>%
          group_by(StudyID, OutcomeKey) %>%
          summarise(
            ctrl_yi = mean(yi, na.rm = TRUE),
            ctrl_vi = mean(vi, na.rm = TRUE),
            .groups = "drop"
          )
        
        # Join treatment to control
        merged_data <- treatment_combos %>%
          left_join(control_combos, by = c("StudyID", "OutcomeKey"))
        
        # Remove rows without matching control
        merged_data <- merged_data %>%
          filter(!is.na(ctrl_yi))
        
        if(nrow(merged_data) == 0) {
          showNotification("No matching pairs found between treatment and control groups", 
                          type = "error", duration = 10)
          return(NULL)
        }
        
        # Calculate differences
        merged_data$yi_diff <- merged_data$yi - merged_data$ctrl_yi
        merged_data$vi_diff <- merged_data$vi + merged_data$ctrl_vi
        merged_data$se_diff <- sqrt(merged_data$vi_diff)
        merged_data$ci_lower <- merged_data$yi_diff - 1.96 * merged_data$se_diff
        merged_data$ci_upper <- merged_data$yi_diff + 1.96 * merged_data$se_diff
        
        # Build labels - include outcome details to differentiate duplicates
        merged_data$label <- sapply(1:nrow(merged_data), function(i) {
          row <- merged_data[i, ]
          
          # Base label: StudyID
          label_parts <- row$StudyID
          
          # Add training type if using subgroup split
          if(use_treatment_subgroup && !is.na(row$TreatmentSubgroup)) {
            label_parts <- paste0(label_parts, "  |  ", row$TreatmentSubgroup)
          }
          
          # Add outcome details if available (ALWAYS include to differentiate duplicates)
          if(!is.na(row$OutcomeLabel) && row$OutcomeLabel != "") {
            label_parts <- paste0(label_parts, "  |  ", row$OutcomeLabel)
          }
          
          label_parts
        })
        
        # Create final data frame
        data_wide <- data.frame(
          StudyID = merged_data$StudyID,
          label = merged_data$label,
          yi = merged_data$yi_diff,
          vi = merged_data$vi_diff,
          se = merged_data$se_diff,
          ci_lower = merged_data$ci_lower,
          ci_upper = merged_data$ci_upper,
          OutcomeKey = merged_data$OutcomeKey,
          stringsAsFactors = FALSE
        )
        
        # Store match_vars for header construction
        if(length(match_vars) > 0) {
          data_wide$match_vars_used <- paste(match_vars, collapse = "|")
        } else {
          data_wide$match_vars_used <- ""
        }
        
        # Add TreatmentSubgroup column for grouping/ordering
        if(use_treatment_subgroup) {
          data_wide$TreatmentSubgroup <- merged_data$TreatmentSubgroup
        }
        
        # Add IDs for three-level model
        data_wide$EffectSizeID <- 1:nrow(data_wide)
        data_wide$StudyGroup <- data_wide$StudyID
        
        # Determine if three-level model is needed
        es_per_study <- table(data_wide$StudyID)
        n_studies <- length(es_per_study)
        n_effects <- nrow(data_wide)
        n_studies_multiple <- sum(es_per_study > 1)
        use_three_level <- n_effects > n_studies && n_studies_multiple > 0
        
        test_arg <- ifelse(input$test_type == "knha", "knha", "z")
        
        if(use_three_level) {
          model <- rma.mv(
            yi = yi, 
            V = vi,
            random = ~ 1 | StudyGroup/EffectSizeID,
            slab = label,  # Store labels for funnel plot
            data = data_wide,
            method = input$ma_model,
            test = ifelse(input$test_type == "knha", "t", "z")
          )
        } else {
          model <- rma(yi = yi, vi = vi, 
                      slab = label,  # Store labels for funnel plot
                      data = data_wide, 
                      method = input$ma_model, test = test_arg)
        }
        
        plot_data <- data_wide
        
        # Build subtitle
        subtitle_parts <- paste0("Difference: ", group1, " - ", group2)
        if(use_treatment_subgroup) {
          n_subgroups <- length(unique(na.omit(data_wide$TreatmentSubgroup)))
          subtitle_parts <- paste0(subtitle_parts, " (split by ", treatment_subgroup_var, ": ", n_subgroups, " types)")
        }
        subtitle_parts <- paste0(subtitle_parts, " [", n_effects, " ES from ", n_studies, " studies]")
        if(use_three_level) {
          subtitle_parts <- paste0(subtitle_parts, " [3-level model]")
        }
        
        list(data = plot_data, model = model,
             plot_title = "Forest Plot: Between-Group Differences",
             subtitle = subtitle_parts)
        
      } else {
        # ============ ONE DIFFERENCE PER STUDY (AVERAGED) ============
        # Original behaviour - average across outcomes within study
        
        # Filter to only these two groups and create pairing
        data <- values$es_data %>%
          filter(Group %in% c(group1, group2)) %>%
          select(StudyID, Group, yi, vi, se, n) %>%
          arrange(StudyID, Group)
        
        # For each study, find if we have both groups
        data_wide <- data %>%
          pivot_wider(
            id_cols = StudyID,
            names_from = Group,
            values_from = c(yi, vi, n),
            values_fn = list(yi = mean, vi = mean, n = mean)  # Average if multiple measurements
          )
        
        # Check if we have both groups
        col_group1_yi <- paste0("yi_", group1)
        col_group2_yi <- paste0("yi_", group2)
        col_group1_vi <- paste0("vi_", group1)
        col_group2_vi <- paste0("vi_", group2)
        
        if(!(col_group1_yi %in% names(data_wide)) || !(col_group2_yi %in% names(data_wide))) {
          showNotification("Missing data for one or both groups", type = "error", duration = 10)
          return(NULL)
        }
        
        # Keep only studies with both groups
        data_wide <- data_wide %>%
          filter(!is.na(.data[[col_group1_yi]]) & !is.na(.data[[col_group2_yi]]))
        
        if(nrow(data_wide) == 0) {
          showNotification("No studies have both groups for pairing", type = "error", duration = 10)
          return(NULL)
        }
        
        # Calculate difference (group1 - group2, e.g., Training - Control)
        data_wide$yi_diff <- data_wide[[col_group1_yi]] - data_wide[[col_group2_yi]]
        data_wide$vi_diff <- data_wide[[col_group1_vi]] + data_wide[[col_group2_vi]]
        data_wide$se_diff <- sqrt(data_wide$vi_diff)
        data_wide$ci_lower <- data_wide$yi_diff - 1.96 * data_wide$se_diff
        data_wide$ci_upper <- data_wide$yi_diff + 1.96 * data_wide$se_diff
        
        # Build labels
        data_wide$label <- paste0(data_wide$StudyID, " (", group1, " - ", group2, ")")
        
        test_arg <- ifelse(input$test_type == "knha", "knha", "z")
        model <- rma(yi = yi_diff, vi = vi_diff, 
                    slab = label,  # Store labels for funnel plot
                    data = data_wide, 
                    method = input$ma_model, test = test_arg)
        
        plot_data <- data.frame(
          StudyID = data_wide$StudyID,
          label = data_wide$label,
          yi = data_wide$yi_diff,
          vi = data_wide$vi_diff,
          se = data_wide$se_diff,
          ci_lower = data_wide$ci_lower,
          ci_upper = data_wide$ci_upper,
          match_vars_used = ""  # No match vars when averaging
        )
        
        list(data = plot_data, model = model,
             plot_title = "Forest Plot: Between-Group Differences",
             subtitle = paste0("Difference: ", group1, " - ", group2, " (averaged within studies)"))
      }
    }
  })
  
  # =================== ENHANCED FOREST PLOT ==================================
  
  forest_plot <- reactive({
    req(values$es_data)
    
    forest_pkg <- prepare_forest_data()
    if (is.null(forest_pkg)) return(NULL)
    
    data  <- forest_pkg$data
    model <- forest_pkg$model
    values$forest_model <- model   # keep this so summary/heterogeneity boxes work
    values$ma_model <- model       # also set this for funnel plot compatibility
    
    # ---- GROUPING LOGIC -----------------------------------------------------
    values$forest_group_models <- NULL  # reset each time
    group_var <- input$forest_group_var
    
    # Check if grouping variable exists in the data
    # For between-group differences, TreatmentSubgroup is the most useful grouping variable
    grouping_enabled <- !is.null(group_var) &&
      nzchar(group_var) &&
      group_var %in% names(data)
    
    # Special case: if user selected a variable that doesn't exist in diff data, 
    # check if TreatmentSubgroup exists as an alternative
    if(!grouping_enabled && input$forest_data_type == "diff" && 
       "TreatmentSubgroup" %in% names(data) && 
       !is.null(group_var) && nzchar(group_var)) {
      # Suggest using TreatmentSubgroup instead
      showNotification(paste0("'", group_var, "' not available for between-group differences. ",
                             "Try grouping by 'TreatmentSubgroup' if using split treatment."), 
                      type = "warning", duration = 5)
    }
    
    test_arg <- ifelse(input$test_type == "knha", "knha", "z")
    
    if (grouping_enabled) {
      # Run separate meta-analyses per level of the grouping variable
      group_models <- lapply(split(data, data[[group_var]]), function(df) {
        tryCatch({
          # Ensure IDs exist
          if(!"EffectSizeID" %in% names(df)) {
            df$EffectSizeID <- 1:nrow(df)
          }
          if(!"StudyGroup" %in% names(df)) {
            df$StudyGroup <- paste(df$StudyID, df$Group, sep = "_")
          }
          
          # Determine if three-level model is needed for this subgroup
          es_per_sg <- table(df$StudyGroup)
          n_sg <- length(es_per_sg)
          n_eff <- nrow(df)
          n_sg_multiple <- sum(es_per_sg > 1)
          use_three_level <- n_eff > n_sg && n_sg_multiple > 0
          
          if(use_three_level) {
            rma.mv(
              yi = yi, 
              V = vi,
              random = ~ 1 | StudyGroup/EffectSizeID,
              data = df,
              method = input$ma_model,
              test = ifelse(input$test_type == "knha", "t", "z")
            )
          } else {
            rma(
              yi = yi, 
              vi = vi,
              data = df,
              method = input$ma_model,
              test = test_arg
            )
          }
        }, error = function(e) NULL)
      })
      # Drop failed models
      group_models <- group_models[!vapply(group_models, is.null, logical(1))]
      if (length(group_models) > 0) {
        values$forest_group_models <- group_models
      }
    }
    
    # ---- BUILD BASIC PLOT DATA ---------------------------------------------
    plot_data <- data.frame(
      study    = if ("label" %in% names(data)) data$label else data$StudyID,
      yi       = data$yi,
      se       = data$se,
      ci_lower = data$yi - 1.96 * data$se,
      ci_upper = data$yi + 1.96 * data$se,
      weight   = weights(model),
      match_vars_used = if("match_vars_used" %in% names(data)) data$match_vars_used else "",
      stringsAsFactors = FALSE
    )
    
    # Attach grouping variable to plot_data (if used)
    if (grouping_enabled) {
      plot_data$group <- factor(data[[group_var]])
    } else {
      plot_data$group <- NA
    }
    
    # ---- ORDERING LOGIC -----------------------------------------------------
    order_var <- input$forest_order_var
    if (is.null(order_var) || !nzchar(order_var)) order_var <- "yi"
    decreasing <- identical(input$forest_order_dir, "desc")
    
    if (order_var == "yi" || !order_var %in% names(data)) {
      base_vec <- data$yi
    } else {
      base_vec <- data[[order_var]]
    }
    
    ord <- order(base_vec, decreasing = decreasing, na.last = TRUE)
    plot_data <- plot_data[ord, , drop = FALSE]
    
    # Recompute row positions after ordering
    plot_data$row <- seq(nrow(plot_data), 1) * input$row_spacing
    
    # ---- Overall diamond vertical position & y-limits -----------------------
    diamond_center_y <- -input$row_spacing * 0.35
    diamond_height   <- input$row_spacing * 0.4
    
    y_max <- max(plot_data$row) + input$row_spacing * 2.0  # Increased to fit header
    y_min <- diamond_center_y - diamond_height - input$row_spacing * 0.8  # Extra space for bracket and Effect Size label
    
    # ---- BASE GG PLOT -------------------------------------------------------
    if (grouping_enabled) {
      p <- ggplot(plot_data, aes(x = yi, y = row,
                                 colour = group, fill = group))
    } else {
      p <- ggplot(plot_data, aes(x = yi, y = row))
    }
    
    # CI band
    if (input$show_ci_band) {
      p <- p +
        annotate(
          "rect",
          xmin = model$ci.lb, xmax = model$ci.ub,
          ymin = y_min,        ymax = y_max,
          fill = input$forest_fill_color, alpha = 0.1
        )
    }
    
    # ---- DISTRIBUTIONS ------------------------------------------------------
    if (input$forest_style == "distribution") {
      dist_data_list <- lapply(1:nrow(plot_data), function(i) {
        # Reduce tail length: use ±2.5 SE from the mean instead of extending beyond CI
        # This captures ~98.8% of the normal distribution without excessive tails
        x_seq <- seq(
          plot_data$yi[i] - 2.5 * plot_data$se[i],
          plot_data$yi[i] + 2.5 * plot_data$se[i],
          length.out = 100
        )
        
        y_dens <- dnorm(x_seq, mean = plot_data$yi[i], sd = plot_data$se[i])
        max_dens <- max(y_dens)
        
        height_scale <- input$dist_width * input$row_spacing * 0.4
        y_scaled <- plot_data$row[i] + (y_dens / max_dens * height_scale)
        
        df <- data.frame(
          x      = x_seq,
          y      = y_scaled,
          y_base = plot_data$row[i],
          study  = i
        )
        if (grouping_enabled) df$group <- plot_data$group[i]
        df
      })
      
      dist_data <- do.call(rbind, dist_data_list)
      
      if (input$show_shaded_dist) {
        if (grouping_enabled) {
          p <- p +
            geom_ribbon(
              data = dist_data,
              aes(x = x, ymin = y_base, ymax = y,
                  group = study, fill = group),
              alpha = input$dist_transparency,
              inherit.aes = FALSE
            )
        } else {
          p <- p +
            geom_ribbon(
              data = dist_data,
              aes(x = x, ymin = y_base, ymax = y, group = study),
              fill  = input$forest_fill_color,
              alpha = input$dist_transparency,
              inherit.aes = FALSE
            )
        }
      }
      
      if (grouping_enabled) {
        p <- p +
          geom_line(
            data = dist_data,
            aes(x = x, y = y, group = study, colour = group),
            alpha = 0.7 * input$line_darkness,
            size  = 0.5 * (0.5 + input$ci_line_thickness),
            inherit.aes = FALSE
          )
      } else {
        p <- p +
          geom_line(
            data = dist_data,
            aes(x = x, y = y, group = study),
            color = input$forest_color,
            alpha = 0.7 * input$line_darkness,
            size  = 0.5 * (0.5 + input$ci_line_thickness),
            inherit.aes = FALSE
          )
      }
    }
    
    # ---- GROUP DIAMONDS -----------------------------------------------------
    # always create this so it's safe to reference later
    group_diamond_centers <- numeric(0)  # named numeric vector
    
    if (grouping_enabled &&
        isTRUE(input$show_group_diamonds) &&
        !is.null(values$forest_group_models) &&
        length(values$forest_group_models) > 0) {
      
      # Order group levels to match plotting order where possible
      group_levels <- names(values$forest_group_models)
      plot_group_levels <- levels(plot_data$group)
      if (!is.null(plot_group_levels)) {
        group_levels <- intersect(plot_group_levels, group_levels)
      }
      if (length(group_levels) == 0) {
        group_levels <- names(values$forest_group_models)
      }
      
      group_diamond_list <- list()
      
      # Stack group diamonds below the overall diamond
      n_g <- length(group_levels)
      if (n_g > 0) {
        gap     <- input$row_spacing * 0.8
        start_y <- diamond_center_y - input$row_spacing * 1.2
        
        for (i in seq_along(group_levels)) {
          g  <- group_levels[i]
          gm <- values$forest_group_models[[g]]
          if (is.null(gm)) next
          
          center_y <- start_y - (i - 1) * gap
          height_g <- input$row_spacing * 0.35
          
          # keep everything inside the plotting window
          y_min <- min(y_min, center_y - height_g - input$row_spacing * 0.2)
          
          diamond_df_g <- data.frame(
            x = c(
              gm$ci.lb,
              coef(gm),
              gm$ci.ub,
              coef(gm),
              gm$ci.lb
            ),
            y = c(
              center_y,
              center_y + height_g,
              center_y,
              center_y - height_g,
              center_y
            ),
            group = factor(g, levels = levels(plot_data$group))
          )
          
          group_diamond_list[[g]]  <- diamond_df_g
          group_diamond_centers[g] <- center_y
        }
      }
      
      if (length(group_diamond_list) > 0) {
        group_diamonds <- do.call(rbind, group_diamond_list)
        
        p <- p +
          geom_polygon(
            data  = group_diamonds,
            aes(x = x, y = y, colour = group, fill = group),
            alpha = 0.25,
            size  = input$point_outline_thickness,
            inherit.aes = FALSE
          )
      }
    }
    
    # ---- POINTS & CIs -------------------------------------------------------
    if (grouping_enabled) {
      p <- p +
        geom_errorbarh(
          aes(xmin = ci_lower, xmax = ci_upper, colour = group),
          height = 0,
          size   = input$ci_line_thickness,
          alpha  = input$line_darkness
        ) +
        geom_point(
          aes(size = weight, colour = group, fill = group),
          shape  = 21,
          stroke = input$point_outline_thickness,
          alpha  = input$line_darkness
        )
    } else {
      p <- p +
        geom_errorbarh(
          aes(xmin = ci_lower, xmax = ci_upper),
          height = 0,
          color  = "black",
          size   = input$ci_line_thickness,
          alpha  = input$line_darkness
        ) +
        geom_point(
          aes(size = weight),
          color  = input$forest_color,
          shape  = 22,
          fill   = input$forest_color,
          stroke = input$point_outline_thickness,
          alpha  = input$line_darkness
        )
    }
    
    p <- p + scale_size_continuous(range = c(2, 6), guide = "none")
    
    if (grouping_enabled) {
      p <- p +
        scale_colour_viridis_d(end = 0.9) +
        scale_fill_viridis_d(end = 0.9)
    }
    
    # ---- OVERALL DIAMOND ----------------------------------------------------
    if (input$show_overall_diamond) {
      diamond_df <- data.frame(
        x = c(
          model$ci.lb,
          coef(model),
          model$ci.ub,
          coef(model),
          model$ci.lb
        ),
        y = c(
          diamond_center_y,
          diamond_center_y + diamond_height,
          diamond_center_y,
          diamond_center_y - diamond_height,
          diamond_center_y
        )
      )
      
      p <- p +
        geom_polygon(
          data  = diamond_df,
          aes(x = x, y = y),
          fill  = input$forest_fill_color,
          color = "black",
          size  = input$point_outline_thickness,
          alpha = input$line_darkness * 0.75,
          inherit.aes = FALSE
        )
    }
    
    # ---- REFERENCE LINES ----------------------------------------------------
    p <- p +
      geom_vline(
        xintercept = 0,
        linetype   = "dotted",
        alpha      = 0.5 * input$line_darkness,
        size       = 0.5
      ) +
      geom_vline(
        xintercept = coef(model),
        color      = input$forest_color,
        size       = 1.2,
        alpha      = 0.6 * input$line_darkness
      )
    
    # ---- TEXT STATS COLUMN --------------------------------------------------
    if (input$show_effect_values || input$show_ci_values || input$show_weights) {
      # Calculate the full data range including group diamonds if present
      data_max_for_stats <- max(plot_data$ci_upper)
      
      # Include group diamond CIs in the calculation if they exist
      if (grouping_enabled &&
          isTRUE(input$show_group_diamonds) &&
          !is.null(values$forest_group_models) &&
          length(values$forest_group_models) > 0) {
        
        group_cis_upper <- sapply(values$forest_group_models, function(gm) {
          if(is.null(gm)) return(NA)
          gm$ci.ub
        })
        
        if(any(!is.na(group_cis_upper))) {
          data_max_for_stats <- max(data_max_for_stats, group_cis_upper, na.rm = TRUE)
        }
      }
      
      # Also include overall diamond if shown
      if (input$show_overall_diamond) {
        data_max_for_stats <- max(data_max_for_stats, model$ci.ub)
      }
      
      x_range    <- diff(range(c(plot_data$ci_lower, plot_data$ci_upper)))
      plot_x_max <- data_max_for_stats + x_range * 0.15  # Increased buffer
      stat_x_pos <- plot_x_max + x_range * 0.6  # Increased spacing from 0.4 to 0.6
      
      # Header will be added later at header_y position (same row as "Study" header)
      # Store stat_x_pos for use in header section
      stats_header_x <- stat_x_pos
      
      for (i in 1:nrow(plot_data)) {
        label_parts <- c()
        
        if (input$show_effect_values) {
          label_parts <- c(label_parts, sprintf("%6.2f", plot_data$yi[i]))
        }
        if (input$show_ci_values) {
          label_parts <- c(
            label_parts,
            sprintf("[%6.2f,%6.2f]", plot_data$ci_lower[i], plot_data$ci_upper[i])
          )
        }
        if (input$show_weights) {
          label_parts <- c(label_parts, sprintf("w=%5.1f%%", plot_data$weight[i]))
        }
        
        if (length(label_parts) > 0) {
          p <- p +
            annotate(
              "text",
              x = stat_x_pos,
              y = plot_data$row[i],
              label = paste(label_parts, collapse = " "),
              hjust  = 0,
              size   = input$forest_text_size * 0.28,
              color  = "black",
              family = "mono"
            )
        }
      }
      
      if (input$show_overall_diamond) {
        overall_label_parts <- c()
        
        if (input$show_effect_values) {
          overall_label_parts <- c(overall_label_parts, sprintf("%6.2f", coef(model)))
        }
        if (input$show_ci_values) {
          overall_label_parts <- c(
            overall_label_parts,
            sprintf("[%6.2f,%6.2f]", model$ci.lb, model$ci.ub)
          )
        }
        if (input$show_overall_pvalue) {
          p_val <- model$pval
          p_text <- if (p_val < 0.001) {
            "p<0.001"
          } else if (p_val < 0.01) {
            sprintf("p=%.3f", p_val)
          } else {
            sprintf("p=%.2f", p_val)
          }
          overall_label_parts <- c(overall_label_parts, p_text)
        }
        
        if (length(overall_label_parts) > 0) {
          label_text <- paste(overall_label_parts, collapse = " ")
          
          # Simple black text (no background box)
          p <- p +
            annotate(
              "text",
              x = stat_x_pos,
              y = diamond_center_y,
              label = label_text,
              hjust  = 0,
              size   = input$forest_text_size * 0.32,
              fontface = "bold",
              color    = "black",  # Black for consistency with groups
              family   = "mono"
            )
        }
      }
      
      # Add color-coded group statistics (if group diamonds are shown)
      if (grouping_enabled &&
          isTRUE(input$show_group_diamonds) &&
          !is.null(values$forest_group_models) &&
          length(values$forest_group_models) > 0 &&
          !is.null(group_diamond_centers) &&
          length(group_diamond_centers) > 0) {
        
        # Get group levels in same order as diamonds
        group_levels <- names(group_diamond_centers)
        
        # Get colors for each group (from viridis palette) - match the diamond colors
        n_groups <- length(group_levels)
        group_colors <- viridis::viridis(n_groups, end = 0.9)
        names(group_colors) <- group_levels
        
        for (i in seq_along(group_levels)) {
          g <- group_levels[i]
          gm <- values$forest_group_models[[g]]
          if (is.null(gm)) next
          
          center_y <- group_diamond_centers[g]
          group_label_parts <- c()
          
          if (input$show_effect_values) {
            group_label_parts <- c(group_label_parts, sprintf("%6.2f", coef(gm)))
          }
          if (input$show_ci_values) {
            group_label_parts <- c(group_label_parts, 
                                  sprintf("[%6.2f,%6.2f]", gm$ci.lb, gm$ci.ub))
          }
          if (input$show_overall_pvalue) {
            p_val <- gm$pval
            p_text <- if (p_val < 0.001) {
              "p<0.001"
            } else if (p_val < 0.01) {
              sprintf("p=%.3f", p_val)
            } else {
              sprintf("p=%.2f", p_val)
            }
            group_label_parts <- c(group_label_parts, p_text)
          }
          
          if (length(group_label_parts) > 0) {
            label_text <- paste(group_label_parts, collapse = " ")
            
            # Simple black text (no background boxes)
            p <- p +
              annotate(
                "text",
                x = stat_x_pos,
                y = center_y,
                label = label_text,
                hjust = 0,
                size = input$forest_text_size * 0.32,  # Match overall font size
                fontface = "bold",
                color = "black",  # Black for easy readability
                family = "mono"
              )
          }
        }
      }
      
      x_plot_min <- min(plot_data$ci_lower, model$ci.lb) - x_range * 0.1
      x_plot_max <- stat_x_pos + x_range * 1.0  # Increased from 0.7 to 1.0 for more text space
    } else {
      x_range    <- diff(range(c(plot_data$ci_lower, plot_data$ci_upper)))
      x_plot_min <- min(plot_data$ci_lower, model$ci.lb) - x_range * 0.1
      x_plot_max <- max(plot_data$ci_upper, model$ci.ub) + x_range * 0.1
    }
    
    # ---- CALCULATE NICE AXIS BREAKS FOR BRACKET --------------------------------
    # Find the data range for the bracket (including group diamonds if shown)
    data_min <- min(plot_data$ci_lower, model$ci.lb)
    data_max <- max(plot_data$ci_upper, model$ci.ub)
    
    # Include group diamond CIs if they exist
    if (grouping_enabled &&
        isTRUE(input$show_group_diamonds) &&
        !is.null(values$forest_group_models) &&
        length(values$forest_group_models) > 0) {
      
      group_cis <- sapply(values$forest_group_models, function(gm) {
        if(is.null(gm)) return(c(NA, NA))
        c(gm$ci.lb, gm$ci.ub)
      })
      
      if(any(!is.na(group_cis))) {
        data_min <- min(data_min, group_cis, na.rm = TRUE)
        data_max <- max(data_max, group_cis, na.rm = TRUE)
      }
    }
    
    # NOW update x_plot_min and x_plot_max to include group diamonds
    # This ensures coord_cartesian doesn't clip them
    x_range_full <- data_max - data_min
    if (input$show_effect_values || input$show_ci_values || input$show_weights) {
      # Stats column shown - keep existing x_plot_max for stats
      x_plot_min <- data_min - x_range_full * 0.1
      # x_plot_max already set for stats column
    } else {
      # No stats column
      x_plot_min <- data_min - x_range_full * 0.1
      x_plot_max <- data_max + x_range_full * 0.1
    }
    
    # Round to nice numbers (0.5 increments for small ranges, 1.0 for larger)
    data_span <- data_max - data_min
    if(data_span <= 2) {
      # Use 0.5 increments
      bracket_min <- floor(data_min * 2) / 2
      bracket_max <- ceiling(data_max * 2) / 2
    } else if(data_span <= 5) {
      # Use 1.0 increments
      bracket_min <- floor(data_min)
      bracket_max <- ceiling(data_max)
    } else {
      # Use 2.0 increments for very large ranges
      bracket_min <- floor(data_min / 2) * 2
      bracket_max <- ceiling(data_max / 2) * 2
    }
    
    # Ensure 0 is included if data spans across it
    if(data_min < 0 && data_max > 0) {
      # Keep the range but ensure nice bounds
      bracket_min <- min(bracket_min, floor(data_min))
      bracket_max <- max(bracket_max, ceiling(data_max))
    }
    
    # ---- COLUMN HEADERS --------------------------------------------------------
    # Add column headers above study labels
    header_y <- max(plot_data$row) + input$row_spacing * 1.2
    
    # Check if TreatmentSubgroup is being used (split treatment by)
    has_treatment_subgroup <- "TreatmentSubgroup" %in% names(data) && 
                              any(!is.na(data$TreatmentSubgroup))
    
    # Check if OutcomeLabel has meaningful content (for header)
    has_outcome_details <- "OutcomeLabel" %in% names(data) &&
                          any(!is.na(data$OutcomeLabel) & data$OutcomeLabel != "" & data$OutcomeLabel != "All")
    
    # Store for later use in y-axis setup
    # Headers will be added via y-axis labels and theme adjustments
    
    # Add "Effect [95% CI] Weight" header if stats are shown - position at top of stats column
    if (exists("stats_header_x")) {
      stats_header_y <- max(plot_data$row) + input$row_spacing * 0.8
      p <- p +
        annotate(
          "text",
          x = stats_header_x,
          y = stats_header_y,
          label = "Effect [95% CI] Weight",
          hjust = 0,
          size = input$forest_text_size * 0.32,
          fontface = "bold",
          color = "black",
          family = "mono"
        )
    }
    
    # ---- BRACKET-STYLE X-AXIS LINE ---------------------------------------------
    # Calculate bracket position (just above the overall diamond area)
    bracket_y <- y_min + input$row_spacing * 0.25
    bracket_height <- input$row_spacing * 0.15
    
    # Create bracket line data using nice axis range
    p <- p +
      # Main horizontal line
      annotate(
        "segment",
        x = bracket_min, xend = bracket_max,
        y = bracket_y, yend = bracket_y,
        color = "black",
        size = 0.5
      ) +
      # Left bracket end (pointing down)
      annotate(
        "segment",
        x = bracket_min, xend = bracket_min,
        y = bracket_y, yend = bracket_y - bracket_height,
        color = "black",
        size = 0.5
      ) +
      # Right bracket end (pointing down)
      annotate(
        "segment",
        x = bracket_max, xend = bracket_max,
        y = bracket_y, yend = bracket_y - bracket_height,
        color = "black",
        size = 0.5
      )
    
    # ---- Y-AXIS LABELS ------------------------------------------------------
    if (input$show_overall_diamond) {
      y_breaks <- c(plot_data$row, diamond_center_y)
      y_labels <- c(plot_data$study, "Overall")
    } else {
      y_breaks <- plot_data$row
      y_labels <- plot_data$study
    }
    
    # Add group-diamond rows under "Overall" if used
    if (grouping_enabled &&
        isTRUE(input$show_group_diamonds) &&
        !is.null(group_diamond_centers) &&
        length(group_diamond_centers) > 0) {
      
      y_breaks <- c(y_breaks, unname(group_diamond_centers))
      y_labels <- c(
        y_labels,
        paste0("Group: ", names(group_diamond_centers))
      )
    }
    
    # Add header row at the top of y-axis labels
    header_y_pos <- max(plot_data$row) + input$row_spacing * 0.9
    y_breaks <- c(header_y_pos, y_breaks)
    
    # Create header label - title case, dynamic based on actual label structure
    
    # Helper function to convert to title case with special handling
    to_title_case <- function(text) {
      # Replace underscores with spaces
      text <- gsub("_", " ", text)
      # Capitalize first letter of each word
      text <- tools::toTitleCase(text)
      text
    }
    
    # Build header dynamically
    header_parts <- c("Study")
    
    # For arm-level, add Group column
    if(input$forest_data_type == "arm") {
      header_parts <- c(header_parts, "Group")
    }
    
    # Add Training Type if present (for between-group differences with split)
    if(has_treatment_subgroup && input$forest_data_type == "diff") {
      header_parts <- c(header_parts, "Training Type")
    }
    
    # Add outcome column headers if present
    if("match_vars_used" %in% names(data)) {
      # Extract match vars from data
      match_vars_str <- unique(data$match_vars_used)[1]
      if(!is.na(match_vars_str) && match_vars_str != "" && 
         !is.null(match_vars_str) && nchar(match_vars_str) > 0) {
        # Split and convert to title case
        match_var_names <- strsplit(match_vars_str, "\\|")[[1]]
        match_var_labels <- sapply(match_var_names, to_title_case)
        header_parts <- c(header_parts, match_var_labels)
      }
    }
    
    # Combine with pipe separators
    header_label <- paste(header_parts, collapse = "  |  ")
    
    y_labels <- c(header_label, y_labels)
    
    # Calculate nice x-axis breaks that match the bracket
    x_axis_breaks <- seq(bracket_min, bracket_max, by = if(bracket_max - bracket_min <= 3) 0.5 else 1)
    
    # Determine x-axis title position - center it on the bracket
    bracket_center <- (bracket_min + bracket_max) / 2
    
    p <- p +
      scale_y_continuous(
        breaks = y_breaks,
        labels = y_labels,
        limits = c(y_min, y_max),
        expand = expansion(mult = c(0.02, 0.02))
      ) +
      scale_x_continuous(
        breaks = x_axis_breaks,
        position = "bottom"
      ) +
      coord_cartesian(xlim = c(bracket_min - (bracket_max - bracket_min) * 0.1, 
                               x_plot_max), 
                      clip = "off") +
      labs(
        x = "Effect Size",
        y = NULL,
        title    = forest_pkg$plot_title,
        subtitle = forest_pkg$subtitle
      ) +
      theme_minimal(base_size = input$forest_text_size) +
      theme(
        text = element_text(color = "black"),  # Override all text to black
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        plot.title       = element_text(face = "bold", color = "black"),
        plot.subtitle    = element_text(color = "black"),
        axis.text.y      = element_text(hjust = 0, size = input$forest_text_size * 0.75, color = "black"),
        axis.text.x      = element_text(color = "black"),
        axis.title.x     = element_text(hjust = (bracket_center - x_plot_min) / (x_plot_max - x_plot_min), color = "black"),
        axis.title.y     = element_text(color = "black"),
        plot.margin      = margin(10, 150, 10, 180),  # Significantly increased left margin
        plot.background  = element_rect(fill = "white", color = NA),
        panel.background = element_rect(fill = "white", color = NA),
        legend.position  = "bottom",
        legend.title     = element_text(face = "bold", size = input$forest_text_size * 0.9, color = "black"),
        legend.text      = element_text(size = input$forest_text_size * 0.8, color = "black"),
        legend.box.background = element_rect(fill = "white", color = "gray80"),
        legend.margin    = margin(5, 5, 5, 5)
      ) +
      guides(
        colour = guide_legend(title = "Group", nrow = 1),
        fill = guide_legend(title = "Group", nrow = 1),
        size = "none"  # Hide the size legend (weight)
      )
    
    p
  })
  
  output$enhanced_forest <- renderPlot({
    # Only update when the button is clicked
    input$update_forest
    
    p <- isolate(forest_plot())
    if (is.null(p)) return(NULL)
    print(p)
  },
  height = function() input$forest_height * 72,
  width  = function() input$forest_width  * 72)
  
  # Summary stats
  output$forest_summary_stats <- renderPrint({
    req(values$forest_model)
    model <- values$forest_model
    
    cat("OVERALL EFFECT\n")
    cat("==============\n")
    cat("k =", model$k, "\n")
    cat("Estimate:", round(coef(model), 3), "\n")
    cat("SE:", round(model$se, 3), "\n")
    cat("95% CI: [", round(model$ci.lb, 3), ",", round(model$ci.ub, 3), "]\n")
    
    # Handle test statistic for both model types
    if(!is.null(model$zval)) {
      cat("z =", round(model$zval, 3), "\n")
    } else if(!is.null(model$QM)) {
      # For rma.mv with t-test
      cat("t =", round(sqrt(model$QM), 3), "\n")
    }
    cat("p =", format.pval(model$pval, digits = 3), "\n")
    
    # ---- Optional group-specific summaries ---------------------------------
    if (isTRUE(input$show_group_summaries) &&
        !is.null(values$forest_group_models) &&
        length(values$forest_group_models) > 0 &&
        !is.null(input$forest_group_var) &&
        nzchar(input$forest_group_var)) {
      
      cat("\nGROUP-SPECIFIC EFFECTS (", input$forest_group_var, ")\n", sep = "")
      cat("=====================================\n")
      
      for (g in names(values$forest_group_models)) {
        gm <- values$forest_group_models[[g]]
        if (is.null(gm)) next
        
        cat("\nGroup:", g, "\n")
        cat("k =", gm$k, "\n")
        cat("Estimate:", round(coef(gm), 3), "\n")
        cat("SE:", round(gm$se, 3), "\n")
        cat("95% CI: [", round(gm$ci.lb, 3), ",", round(gm$ci.ub, 3), "]\n")
        
        # Handle test statistic for both model types in subgroups
        if(!is.null(gm$zval)) {
          cat("z =", round(gm$zval, 3), "\n")
        } else if(!is.null(gm$QM)) {
          cat("t =", round(sqrt(gm$QM), 3), "\n")
        }
        cat("p =", format.pval(gm$pval, digits = 3), "\n")
      }
    }
  })
  
  output$forest_heterogeneity_stats <- renderPrint({
    req(values$forest_model)
    model <- values$forest_model
    
    cat("HETEROGENEITY\n")
    cat("=============\n")
    
    if(inherits(model, "rma.mv")) {
      # Three-level model - extract variance components
      sigma2_level3 <- model$sigma2[1]  # Between study-groups
      sigma2_level2 <- model$sigma2[2]  # Within study-groups
      total_var <- sigma2_level3 + sigma2_level2
      
      cat("Level 3 σ² (between) =", round(sigma2_level3, 4), "\n")
      cat("Level 2 σ² (within) =", round(sigma2_level2, 4), "\n")
      cat("Total τ² =", round(total_var, 4), "\n")
      cat("τ =", round(sqrt(total_var), 4), "\n\n")
      
      # Calculate I² for three-level model
      # Need effect size data for typical sampling variance
      if(!is.null(values$es_data)) {
        typical_vi <- mean(values$es_data$vi, na.rm = TRUE)
        I2_total <- 100 * total_var / (total_var + typical_vi)
        I2_level3 <- 100 * sigma2_level3 / (total_var + typical_vi)
        I2_level2 <- 100 * sigma2_level2 / (total_var + typical_vi)
        
        cat("I² (total) =", round(I2_total, 1), "%\n")
        cat("I² (level 3) =", round(I2_level3, 1), "%\n")
        cat("I² (level 2) =", round(I2_level2, 1), "%\n\n")
      }
      
      cat("Q =", round(model$QE, 2), "\n")
      cat("df =", model$QEdf, "\n")
      cat("p =", format.pval(model$QEp, digits = 3), "\n")
      
    } else {
      # Standard two-level model
      cat("τ² =", round(model$tau2, 4), "\n")
      cat("τ =", round(sqrt(model$tau2), 4), "\n")
      cat("I² =", round(model$I2, 1), "%\n")
      cat("H² =", round(model$H2, 2), "\n\n")
      cat("Q =", round(model$QE, 2), "\n")
      cat("df =", model$k - model$p, "\n")
      cat("p =", format.pval(model$QEp, digits = 3), "\n")
    }
  })
  
  # =================== FUNNEL PLOT (FIXED) ===================================
  
  output$funnel_plot <- renderPlot({
    # Trigger on update button
    input$update_funnel
    
    # Use whichever model is available (forest_model or ma_model)
    model_to_use <- if(!is.null(values$forest_model)) {
      values$forest_model
    } else if(!is.null(values$ma_model)) {
      values$ma_model
    } else {
      NULL
    }
    
    req(model_to_use, values$es_data)
    
    # Use isolate to prevent reactive dependencies
    isolate({
      # Create funnel plot
      if(input$contour_enhanced) {
        funnel(model_to_use, level = c(90, 95, 99), 
              shade = c("white", "gray75", "gray60"), legend = TRUE)
      } else {
        funnel(model_to_use)
      }
      
      # Add labels - use data from the MODEL, not values$es_data
      if(input$label_funnel == TRUE) {
        # Extract data from the fitted model
        model_yi <- model_to_use$yi
        model_vi <- model_to_use$vi
        model_se <- sqrt(model_vi)
        
        # Get study labels from model data - check multiple sources
        if(!is.null(model_to_use$data)) {
          # Try to find label or StudyID column in model's data
          if("label" %in% names(model_to_use$data)) {
            labels <- model_to_use$data$label
          } else if("StudyID" %in% names(model_to_use$data)) {
            labels <- model_to_use$data$StudyID
          } else if(!is.null(model_to_use$slab)) {
            labels <- model_to_use$slab
          } else {
            labels <- paste("Effect", seq_along(model_yi))
          }
        } else if(!is.null(model_to_use$slab)) {
          labels <- model_to_use$slab
        } else {
          labels <- paste("Effect", seq_along(model_yi))
        }
        
        # Only label if not too many
        if(length(labels) <= input$max_funnel_labels) {
          text(x = model_yi, 
               y = model_se,
               labels = labels, 
               cex = input$label_size, 
               pos = 3,
               col = "black")
        }
      }
    })
  })
  
  output$egger_results <- renderPrint({
    req(input$run_egger)
    
    # Use whichever model is available
    model_to_use <- if(!is.null(values$forest_model)) {
      values$forest_model
    } else if(!is.null(values$ma_model)) {
      values$ma_model
    } else {
      NULL
    }
    
    req(model_to_use)
    
    if(inherits(model_to_use, "rma.mv")) {
      # For three-level models, aggregate to one ES per study-group for bias tests
      cat("Note: Using aggregated effect sizes for publication bias test\n")
      cat("(One effect size per study-group)\n\n")
      
      # Aggregate by taking the mean effect size per StudyGroup
      agg_data <- values$es_data %>%
        group_by(StudyGroup) %>%
        summarise(
          yi_agg = mean(yi, na.rm = TRUE),
          vi_agg = mean(vi, na.rm = TRUE),  # Conservative approach
          .groups = 'drop'
        )
      
      # Fit simple two-level model for bias test
      temp_model <- tryCatch(
        rma(yi = yi_agg, vi = vi_agg, data = agg_data, method = "REML"),
        error = function(e) NULL
      )
      
      if(is.null(temp_model)) {
        cat("Error: Could not fit aggregated model for bias test.\n")
        return(NULL)
      }
      
      egger <- regtest(temp_model)
      values$egger_pval <- egger$pval
      
    } else {
      # Standard two-level model - use directly
      egger <- regtest(model_to_use)
      values$egger_pval <- egger$pval
    }
    
    cat("Egger's test for funnel plot asymmetry:\n")
    cat("z =", round(egger$zval, 3), "\n")
    cat("p =", format.pval(egger$pval, 3), "\n")
    
    if(egger$pval < 0.05) {
      cat("\n*** Significant asymmetry detected (p < 0.05) ***\n")
      cat("Consider trim-and-fill analysis below.\n")
    } else {
      cat("\nNo significant asymmetry detected.\n")
    }
  })
  
  
  # Output for conditional panel - whether Egger's test is significant
  output$egger_significant <- reactive({
    req(input$run_egger)
    
    # Use whichever model is available
    model_to_use <- if(!is.null(values$forest_model)) {
      values$forest_model
    } else if(!is.null(values$ma_model)) {
      values$ma_model
    } else {
      return(FALSE)
    }
    
    if(inherits(model_to_use, "rma.mv")) {
      # Use aggregated data
      agg_data <- values$es_data %>%
        group_by(StudyGroup) %>%
        summarise(
          yi_agg = mean(yi, na.rm = TRUE),
          vi_agg = mean(vi, na.rm = TRUE),
          .groups = 'drop'
        )
      
      temp_model <- tryCatch(
        rma(yi = yi_agg, vi = vi_agg, data = agg_data, method = "REML"),
        error = function(e) NULL
      )
      
      if(is.null(temp_model)) return(FALSE)
      egger <- regtest(temp_model)
      
    } else {
      egger <- regtest(model_to_use)
    }
    
    return(egger$pval < 0.05)
  })
  outputOptions(output, "egger_significant", suspendWhenHidden = FALSE)
  
  # Trim-and-fill results UI
  output$trimfill_results_ui <- renderUI({
    req(input$run_egger, input$run_trimfill)
    
    # Use whichever model is available
    model_to_use <- if(!is.null(values$forest_model)) {
      values$forest_model
    } else if(!is.null(values$ma_model)) {
      values$ma_model
    } else {
      return(div(class = "alert alert-warning",
        icon("exclamation-triangle"),
        " Please generate a forest plot first to enable publication bias tests."
      ))
    }
    
    # Prepare aggregated data if needed
    if(inherits(model_to_use, "rma.mv")) {
      agg_data <- values$es_data %>%
        group_by(StudyGroup) %>%
        summarise(
          yi_agg = mean(yi, na.rm = TRUE),
          vi_agg = mean(vi, na.rm = TRUE),
          .groups = 'drop'
        )
      
      temp_model <- tryCatch(
        rma(yi = yi_agg, vi = vi_agg, data = agg_data, method = "REML"),
        error = function(e) NULL
      )
      
      if(is.null(temp_model)) {
        return(div(class = "alert alert-danger",
          icon("exclamation-triangle"),
          " Error: Could not fit aggregated model for trim-and-fill."
        ))
      }
      
      egger <- regtest(temp_model)
      model_for_trimfill <- temp_model
      is_aggregated <- TRUE
      
    } else {
      egger <- regtest(model_to_use)
      model_for_trimfill <- model_to_use
      is_aggregated <- FALSE
    }
    
    if(egger$pval >= 0.05) {
      return(
        div(class = "alert alert-info",
          icon("info-circle"),
          " Egger's test is not significant (p >= 0.05).",
          br(),
          "Trim-and-fill analysis is typically only warranted when funnel plot asymmetry is detected."
        )
      )
    }
    
    # Run trim-and-fill
    tryCatch({
      taf <- trimfill(model_for_trimfill, side = input$trimfill_side)
      values$trimfill_model <- taf
      
      # Get original estimate
      orig_est <- coef(model_for_trimfill)
      orig_ci_lb <- model_for_trimfill$ci.lb
      orig_ci_ub <- model_for_trimfill$ci.ub
      
      # Get adjusted estimate
      adj_est <- coef(taf)
      adj_ci_lb <- taf$ci.lb
      adj_ci_ub <- taf$ci.ub
      
      # Number of studies filled
      k_filled <- taf$k0
      
      # Create output
      tagList(
        div(class = "well",
          h4(icon("balance-scale"), " Trim-and-Fill Analysis Results"),
          if(is_aggregated) {
            p(em("Note: Analysis performed on aggregated effect sizes (one per study-group)"),
              style = "color: #0066cc;")
          },
          hr(),
          
          h5("Number of Studies:"),
          p(strong("Original:"), taf$k - k_filled),
          p(strong("Imputed:"), k_filled),
          p(strong("Total (after filling):"), taf$k),
          
          hr(),
          h5("Effect Size Estimates:"),
          p(strong("Original:"), sprintf("%.3f [%.3f, %.3f]", orig_est, orig_ci_lb, orig_ci_ub)),
          p(strong("Adjusted:"), sprintf("%.3f [%.3f, %.3f]", adj_est, adj_ci_lb, adj_ci_ub)),
          p(strong("Change:"), sprintf("%.3f (%.1f%%)", 
                                         adj_est - orig_est,
                                         abs(adj_est - orig_est) / abs(orig_est) * 100)),
          
          hr(),
          h5("Interpretation:"),
          if(k_filled == 0) {
            p(icon("check-circle"), " No studies were imputed. Funnel plot appears symmetric.",
              style = "color: green;")
          } else {
            p(icon("exclamation-triangle"), 
              sprintf(" %d studies were imputed, suggesting potential publication bias.", k_filled),
              style = "color: orange;")
          },
          
          if(abs(adj_est - orig_est) / abs(orig_est) > 0.1) {
            p(icon("exclamation-circle"),
              " The adjusted estimate differs by >10% from the original.",
              br(),
              "Consider the possibility of publication bias in your conclusions.",
              style = "color: orange;")
          } else {
            p(icon("info-circle"),
              " The adjusted estimate is similar to the original (<10% change).",
              br(),
              "Results appear robust to potential publication bias.",
              style = "color: #0066cc;")
          }
        )
      )
      
    }, error = function(e) {
      div(class = "alert alert-danger",
        icon("exclamation-triangle"),
        " Error running trim-and-fill: ", e$message
      )
    })
  })
  
  
  # Trim-and-fill funnel plot
  output$trimfill_funnel_plot <- renderPlot({
    req(values$trimfill_model)
    input$update_funnel
    
    tryCatch({
      # The trimfill_model is already aggregated if we had rma.mv
      funnel(values$trimfill_model, level = c(90, 95, 99),
             shade = c("white", "gray55", "gray75"), 
             refline = 0,
             main = "Trim-and-Fill Adjusted Funnel Plot")
      
      if(inherits(values$ma_model, "rma.mv")) {
        mtext("Note: Based on aggregated effect sizes (one per study-group)", 
              side = 1, line = 4, cex = 0.8, col = "blue")
      }
    }, error = function(e) {
      plot.new()
      text(0.5, 0.5, paste("Error creating plot:", e$message), col = "red")
    })
  })
  

  # =================== METHODS & RESULTS WRITEUP ==============================
  
  # Helper function to format p-values for writeup
  format_p_writeup <- function(p) {
    if(p < 0.001) return("< .001")
    else if(p < 0.01) return(paste0("= ", sprintf("%.3f", p)))
    else return(paste0("= ", sprintf("%.2f", p)))
  }
  
  # Helper function to get model description
  get_model_description <- function(method, test, citation_style = "apa") {
    model_text <- switch(method,
      "REML" = "restricted maximum likelihood (REML)",
      "DL" = "DerSimonian-Laird",
      "FE" = "fixed-effect",
      "random-effects"
    )
    
    test_text <- if(test == "knha") {
      "with Knapp-Hartung adjustment for confidence intervals and tests"
    } else {
      "with standard z-tests"
    }
    
    ref <- if(citation_style == "apa") {
      switch(method,
        "REML" = "(Viechtbauer, 2005)",
        "DL" = "(DerSimonian & Laird, 1986)",
        "FE" = "(Hedges & Olkin, 1985)",
        ""
      )
    } else {
      switch(method,
        "REML" = "[1]",
        "DL" = "[1]",
        "FE" = "[1]",
        ""
      )
    }
    
    list(model = model_text, test = test_text, ref = ref)
  }
  
  # Generate writeup text
  generate_writeup_text <- reactive({
    req(values$ma_model)
    
    is_md <- input$writeup_format %in% c("md", "docx")
    bold_start <- if(is_md) "**" else ""
    bold_end <- if(is_md) "**" else ""
    italic_start <- if(is_md) "*" else ""
    italic_end <- if(is_md) "*" else ""
    heading2 <- if(is_md) "## " else ""
    heading3 <- if(is_md) "### " else ""
    
    sections <- list()
    references <- c()
    
    # ===== METHODS SECTION =====
    if(isTRUE(input$include_methods)) {
      methods_parts <- c()
      
      # Data and filtering
      methods_parts <- c(methods_parts, paste0(heading2, "Methods\n"))
      methods_parts <- c(methods_parts, paste0(heading3, "Data Extraction and Preparation\n"))
      
      n_original <- if(!is.null(values$raw_data)) nrow(values$raw_data) else "N/A"
      n_filtered <- if(!is.null(values$filtered_data)) nrow(values$filtered_data) else nrow(values$es_data)
      n_effects <- nrow(values$es_data)
      
      filter_text <- ""
      if(length(values$filter_history) > 0 && !is.null(values$filtered_data)) {
        filter_details <- sapply(names(values$filter_history), function(var) {
          vals <- values$filter_history[[var]]
          if(length(vals) <= 3) {
            paste0(var, " (", paste(vals, collapse = ", "), ")")
          } else {
            paste0(var, " (", length(vals), " categories)")
          }
        })
        filter_text <- paste0(" Studies were filtered based on the following criteria: ",
                             paste(filter_details, collapse = "; "),
                             ". After applying these inclusion criteria, ", 
                             n_filtered, " observations remained for analysis.")
      } else if(length(values$active_filter_vars) > 0 && !is.null(values$filtered_data)) {
        filter_text <- paste0(" After applying inclusion criteria based on ",
                             paste(values$active_filter_vars, collapse = ", "),
                             ", ", n_filtered, " observations remained for analysis.")
      }
      
      methods_parts <- c(methods_parts, paste0(
        "Data were extracted from the included studies.", filter_text,
        " A total of ", n_effects, " effect sizes were calculated for the meta-analysis.\n\n"
      ))
      
      # Effect size calculation
      methods_parts <- c(methods_parts, paste0(heading3, "Effect Size Calculation\n"))
      
      es_type_text <- switch(input$es_type,
        "SMCR" = "standardized mean change using raw score standardization (SMCR)",
        "SMD" = "standardized mean difference (SMD; Cohen's d)",
        "MD" = "raw mean difference (MD)"
      )
      
      hedges_text <- if(isTRUE(input$apply_hedges)) {
        paste0(" Hedges' ", italic_start, "g", italic_end, 
               " correction was applied to account for small sample bias",
               if(input$citation_style == "apa") " (Hedges, 1981)" else " [2]", ".")
      } else ""
      
      corr_text <- switch(input$corr_source,
        "fixed" = paste0(" A pre-post correlation of ", italic_start, "r", italic_end, 
                        " = ", input$corr_value, " was assumed for variance estimation."),
        "column" = " Pre-post correlations were obtained from the data for variance estimation.",
        "compute" = " Pre-post correlations were computed from the change score standard deviations.",
        ""
      )
      
      methods_parts <- c(methods_parts, paste0(
        "Effect sizes were calculated as ", es_type_text, ".",
        hedges_text, corr_text, "\n\n"
      ))
      
      if(isTRUE(input$apply_hedges)) {
        references <- c(references, 
          if(input$citation_style == "apa") {
            "Hedges, L. V. (1981). Distribution theory for Glass's estimator of effect size and related estimators. Journal of Educational Statistics, 6(2), 107-128."
          } else {
            "Hedges LV. Distribution theory for Glass's estimator of effect size and related estimators. J Educ Stat. 1981;6(2):107-128."
          }
        )
      }
      
      # Statistical analysis
      methods_parts <- c(methods_parts, paste0(heading3, "Statistical Analysis\n"))
      
      model_info <- get_model_description(input$ma_model, input$test_type, input$citation_style)
      
      model_type_text <- if(input$ma_model == "FE") "fixed-effect" else "random-effects"
      
      # Determine if three-level model was used based on stored model type
      is_three_level <- inherits(values$ma_model, "rma.mv")
      
      if(is_three_level) {
        # Three-level model description
        methods_parts <- c(methods_parts, paste0(
          "A three-level ", model_type_text, " meta-analytic model was conducted using the ", 
          model_info$model, " estimator ", model_info$ref, " to account for dependent effect sizes within studies. ",
          "This hierarchical model partitions variance into sampling variance (level 1), within-study variance (level 2), ",
          "and between-study variance (level 3)", 
          if(input$citation_style == "apa") " (Cheung, 2014)" else " [3]", ". ",
          model_info$test, ". Heterogeneity was assessed using Cochran's ", italic_start, "Q", italic_end,
          " statistic and ", italic_start, "I", italic_end, "² indices at each level",
          if(input$citation_style == "apa") " (Higgins & Thompson, 2002)" else " [4]", ".\n\n"
        ))
        
        references <- c(references,
          if(input$citation_style == "apa") {
            c("Cheung, M. W. L. (2014). Modeling dependent effect sizes with three-level meta-analyses: A structural equation modeling approach. Psychological Methods, 19(2), 211-229.",
              "Viechtbauer, W. (2005). Bias and efficiency of meta-analytic variance estimators in the random-effects model. Journal of Educational and Behavioral Statistics, 30(3), 261-293.",
              "Higgins, J. P., & Thompson, S. G. (2002). Quantifying heterogeneity in a meta-analysis. Statistics in Medicine, 21(11), 1539-1558.")
          } else {
            c("Cheung MWL. Modeling dependent effect sizes with three-level meta-analyses: A structural equation modeling approach. Psychol Methods. 2014;19(2):211-229.",
              "Viechtbauer W. Bias and efficiency of meta-analytic variance estimators in the random-effects model. J Educ Behav Stat. 2005;30(3):261-293.",
              "Higgins JP, Thompson SG. Quantifying heterogeneity in a meta-analysis. Stat Med. 2002;21(11):1539-1558.")
          }
        )
      } else {
        # Standard two-level model description
        methods_parts <- c(methods_parts, paste0(
          "A ", model_type_text, " meta-analysis was conducted using the ", 
          model_info$model, " estimator ", model_info$ref, ", ", model_info$test, 
          ". Heterogeneity was assessed using Cochran's ", italic_start, "Q", italic_end,
          " statistic and the ", italic_start, "I", italic_end, "² index",
          if(input$citation_style == "apa") " (Higgins & Thompson, 2002)" else " [3]", ".\n\n"
        ))
        
        references <- c(references,
          if(input$citation_style == "apa") {
            c("Viechtbauer, W. (2005). Bias and efficiency of meta-analytic variance estimators in the random-effects model. Journal of Educational and Behavioral Statistics, 30(3), 261-293.",
              "Higgins, J. P., & Thompson, S. G. (2002). Quantifying heterogeneity in a meta-analysis. Statistics in Medicine, 21(11), 1539-1558.")
          } else {
            c("Viechtbauer W. Bias and efficiency of meta-analytic variance estimators in the random-effects model. J Educ Behav Stat. 2005;30(3):261-293.",
              "Higgins JP, Thompson SG. Quantifying heterogeneity in a meta-analysis. Stat Med. 2002;21(11):1539-1558.")
          }
        )
      }
      
      # Publication bias methods
      if(isTRUE(input$run_egger)) {
        # Handle both rma and rma.mv models
        if(inherits(values$ma_model, "rma.mv")) {
          # Aggregate for Egger test in methods writeup
          agg_data <- values$es_data %>%
            group_by(StudyGroup) %>%
            summarise(
              yi_agg = mean(yi, na.rm = TRUE),
              vi_agg = mean(vi, na.rm = TRUE),
              .groups = 'drop'
            )
          
          temp_model <- tryCatch(
            rma(yi = yi_agg, vi = vi_agg, data = agg_data, method = "REML"),
            error = function(e) NULL
          )
          
          if(!is.null(temp_model)) {
            egger <- regtest(temp_model)
          } else {
            egger <- NULL
          }
        } else {
          egger <- regtest(values$ma_model)
        }
        
        if(!is.null(egger)) {
        
        methods_parts <- c(methods_parts, paste0(heading3, "Publication Bias Assessment\n"))
        
        pub_bias_text <- paste0(
          "Publication bias was assessed through visual inspection of funnel plots and formally tested using Egger's regression test for funnel plot asymmetry",
          if(input$citation_style == "apa") " (Egger et al., 1997)" else " [4]", "."
        )
        
        if(isTRUE(input$run_trimfill) && egger$pval < 0.05 && !is.null(values$trimfill_model)) {
          pub_bias_text <- paste0(pub_bias_text,
            " When significant asymmetry was detected, Duval and Tweedie's trim-and-fill procedure was applied to estimate the number of missing studies and provide bias-adjusted effect estimates",
            if(input$citation_style == "apa") " (Duval & Tweedie, 2000)" else " [5]", "."
          )
          
          references <- c(references,
            if(input$citation_style == "apa") {
              "Duval, S., & Tweedie, R. (2000). Trim and fill: A simple funnel-plot-based method of testing and adjusting for publication bias in meta-analysis. Biometrics, 56(2), 455-463."
            } else {
              "Duval S, Tweedie R. Trim and fill: A simple funnel-plot-based method of testing and adjusting for publication bias in meta-analysis. Biometrics. 2000;56(2):455-463."
            }
          )
        }
        
        }
        methods_parts <- c(methods_parts, paste0(pub_bias_text, "\n\n"))
        
        references <- c(references,
          if(input$citation_style == "apa") {
            "Egger, M., Davey Smith, G., Schneider, M., & Minder, C. (1997). Bias in meta-analysis detected by a simple, graphical test. BMJ, 315(7109), 629-634."
          } else {
            "Egger M, Davey Smith G, Schneider M, Minder C. Bias in meta-analysis detected by a simple, graphical test. BMJ. 1997;315(7109):629-634."
          }
        )
      }
      
      methods_parts <- c(methods_parts, paste0(
        "All analyses were conducted using the metafor package",
        if(input$citation_style == "apa") " (Viechtbauer, 2010)" else " [6]",
        " in R.\n\n"
      ))
      
      references <- c(references,
        if(input$citation_style == "apa") {
          "Viechtbauer, W. (2010). Conducting meta-analyses in R with the metafor package. Journal of Statistical Software, 36(3), 1-48."
        } else {
          "Viechtbauer W. Conducting meta-analyses in R with the metafor package. J Stat Softw. 2010;36(3):1-48."
        }
      )
      
      sections$methods <- paste(methods_parts, collapse = "")
    }
    
    # ===== RESULTS SECTION =====
    if(isTRUE(input$include_results)) {
      results_parts <- c()
      
      results_parts <- c(results_parts, paste0(heading2, "Results\n"))
      results_parts <- c(results_parts, paste0(heading3, "Overall Effect\n"))
      
      model <- values$ma_model
      k <- model$k
      est <- coef(model)
      ci_lb <- model$ci.lb
      ci_ub <- model$ci.ub
      
      # Determine effect size label
      es_label <- if(isTRUE(input$apply_hedges)) "g" else {
        switch(input$es_type, "SMCR" = "d", "SMD" = "d", "MD" = "MD")
      }
      
      # Format test statistic
      # Format test statistic (handle both rma and rma.mv)
      if(input$test_type == "knha") {
        if(inherits(model, "rma.mv")) {
          # For rma.mv, use QM test statistic
          test_stat <- paste0(italic_start, "t", italic_end, " = ", 
                             sprintf("%.2f", sqrt(model$QM)))
        } else {
          test_stat <- paste0(italic_start, "t", italic_end, "(", model$dfs, ") = ", 
                             sprintf("%.2f", model$zval))
        }
      } else {
        test_stat <- paste0(italic_start, "z", italic_end, " = ", 
                           sprintf("%.2f", model$zval))
      }
      
      
      results_parts <- c(results_parts, paste0(
        "The meta-analysis included ", bold_start, k, bold_end, " effect sizes. ",
        "The overall pooled effect was ", bold_start, italic_start, es_label, italic_end,
        " = ", sprintf("%.2f", est), bold_end,
        " (95% CI [", sprintf("%.2f", ci_lb), ", ", sprintf("%.2f", ci_ub), "]), ",
        test_stat, ", ", italic_start, "p", italic_end, " ", format_p_writeup(model$pval),
        if(model$pval < 0.05) ", indicating a statistically significant effect." else 
          ", which was not statistically significant.",
        "\n\n"
      ))
      
      # Heterogeneity
      results_parts <- c(results_parts, paste0(heading3, "Heterogeneity\n"))
      
      # Extract heterogeneity statistics (handle both rma and rma.mv)
      if(inherits(model, "rma.mv")) {
        # For three-level models
        sigma2_level3 <- model$sigma2[1]
        sigma2_level2 <- model$sigma2[2]
        tau2 <- sigma2_level3 + sigma2_level2  # Total variance
        
        # Calculate I² for three-level model
        typical_vi <- mean(values$es_data$vi, na.rm = TRUE)
        I2 <- 100 * tau2 / (tau2 + typical_vi)
        
        Q <- model$QE
        Q_df <- model$QEdf
        Q_p <- model$QEp
      } else {
        # For standard two-level models
        tau2 <- model$tau2
        I2 <- model$I2
        Q <- model$QE
        Q_df <- model$k - model$p
        Q_p <- model$QEp
      }
      
      I2_interp <- if(I2 < 25) "low" else if(I2 < 50) "low to moderate" else 
                   if(I2 < 75) "moderate to substantial" else "substantial"
      
      results_parts <- c(results_parts, paste0(
        "Heterogeneity analysis revealed ", I2_interp, " heterogeneity among effect sizes, ",
        italic_start, "Q", italic_end, "(", Q_df, ") = ", sprintf("%.2f", Q), ", ",
        italic_start, "p", italic_end, " ", format_p_writeup(Q_p), ", ",
        italic_start, "I", italic_end, "² = ", sprintf("%.1f", I2), "%, ",
        "τ² = ", sprintf("%.3f", tau2), ".\n\n"
      ))
      
      # Publication bias results
      if(isTRUE(input$run_egger)) {
        results_parts <- c(results_parts, paste0(heading3, "Publication Bias\n"))
        
        # Handle both rma and rma.mv models for Egger's test
        if(inherits(values$ma_model, "rma.mv")) {
          # Use aggregated data for Egger's test
          agg_data <- values$es_data %>%
            group_by(StudyGroup) %>%
            summarise(
              yi_agg = mean(yi, na.rm = TRUE),
              vi_agg = mean(vi, na.rm = TRUE),
              .groups = 'drop'
            )
          
          temp_model <- tryCatch(
            rma(yi = yi_agg, vi = vi_agg, data = agg_data, method = "REML"),
            error = function(e) NULL
          )
          
          if(!is.null(temp_model)) {
            egger <- regtest(temp_model)
          } else {
            egger <- NULL
          }
        } else {
          egger <- regtest(values$ma_model)
        }
        
        if(is.null(egger)) {
          results_parts <- c(results_parts, "Publication bias assessment was not available.\n\n")
        } else {
        
        egger_result <- paste0(
          "Egger's regression test ",
          if(egger$pval < 0.05) "indicated significant" else "did not indicate significant",
          " funnel plot asymmetry, ", italic_start, "z", italic_end, " = ", 
          sprintf("%.2f", egger$zval), ", ", italic_start, "p", italic_end, " ", 
          format_p_writeup(egger$pval), "."
        )
        
        # Add trim-and-fill results if applicable
        if(isTRUE(input$run_trimfill) && egger$pval < 0.05 && !is.null(values$trimfill_model)) {
          taf <- values$trimfill_model
          k0 <- taf$k0
          adj_est <- coef(taf)
          adj_ci_lb <- taf$ci.lb
          adj_ci_ub <- taf$ci.ub
          
          side_text <- if(input$trimfill_side == "left") "left (smaller effects)" else 
                       "right (larger effects)"
          
          egger_result <- paste0(egger_result, 
            " The trim-and-fill procedure identified ", bold_start, k0, bold_end, 
            " potentially missing studies on the ", side_text, " side of the funnel plot. ",
            "After imputing these studies, the adjusted effect estimate was ",
            italic_start, es_label, italic_end, " = ", sprintf("%.2f", adj_est),
            " (95% CI [", sprintf("%.2f", adj_ci_lb), ", ", sprintf("%.2f", adj_ci_ub), "]), ",
            "compared to the original estimate of ", italic_start, es_label, italic_end, 
            " = ", sprintf("%.2f", est), "."
          )
          
          # Add interpretation
          change_pct <- abs((adj_est - est) / est * 100)
          if(change_pct > 20) {
            egger_result <- paste0(egger_result,
              " This represents a ", sprintf("%.0f", change_pct), 
              "% reduction in the effect size magnitude, suggesting that publication bias may have inflated the original estimate."
            )
          } else if(change_pct > 10) {
            egger_result <- paste0(egger_result,
              " This represents a modest ", sprintf("%.0f", change_pct),
              "% change in the effect size."
            )
          } else {
            egger_result <- paste0(egger_result,
              " The adjusted estimate remained similar to the original, suggesting the findings are relatively robust to potential publication bias."
            )
          }
        } else if(egger$pval >= 0.05) {
          egger_result <- paste0(egger_result,
            " Visual inspection of the funnel plot supported this finding, suggesting low risk of publication bias."
          )
        }
        
        results_parts <- c(results_parts, paste0(egger_result, "\n\n"))
      }
        }
      
      sections$results <- paste(results_parts, collapse = "")
    }
    
    # ===== REFERENCES SECTION =====
    if(isTRUE(input$include_references) && length(references) > 0) {
      references <- unique(references)
      ref_parts <- c(paste0(heading2, "References\n"))
      
      for(i in seq_along(references)) {
        if(input$citation_style == "vancouver") {
          ref_parts <- c(ref_parts, paste0("[", i, "] ", references[i], "\n"))
        } else {
          ref_parts <- c(ref_parts, paste0(references[i], "\n\n"))
        }
      }
      
      sections$references <- paste(ref_parts, collapse = "")
    }
    
    # Combine all sections
    full_text <- paste(c(sections$methods, sections$results, sections$references), collapse = "\n")
    
    return(full_text)
  })
  
  # Generate writeup on button click
  observeEvent(input$generate_writeup, {
    req(values$ma_model)
    values$writeup_text <- generate_writeup_text()
    showNotification("Write-up generated!", type = "message")
  })
  
  # Preview output
  output$writeup_preview <- renderUI({
    req(values$writeup_text)
    
    if(input$writeup_format == "md") {
      # Render markdown
      HTML(paste0("<div style='white-space: pre-wrap; font-family: Georgia, serif;'>", 
                 gsub("\\*\\*(.+?)\\*\\*", "<b>\\1</b>",
                      gsub("\\*(.+?)\\*", "<i>\\1</i>",
                           gsub("## ", "<h3>", 
                                gsub("### ", "<h4>",
                                     gsub("\n", "<br>", values$writeup_text))))),
                 "</div>"))
    } else {
      # Plain text
      pre(style = "white-space: pre-wrap; font-family: monospace;", values$writeup_text)
    }
  })
  
  # Download writeup
  output$download_writeup <- downloadHandler(
    filename = function() {
      ext <- switch(input$writeup_format,
        "txt" = "txt",
        "md" = "md", 
        "docx" = "docx"
      )
      paste0("meta_analysis_writeup_", Sys.Date(), ".", ext)
    },
    content = function(file) {
      req(values$writeup_text)
      
      if(input$writeup_format == "docx") {
        # For Word documents, we need to use a package
        # Check if officer is available, otherwise fall back to txt
        if(requireNamespace("officer", quietly = TRUE)) {
          doc <- officer::read_docx()
          
          # Split text into paragraphs and add formatting
          lines <- strsplit(values$writeup_text, "\n")[[1]]
          
          for(line in lines) {
            if(grepl("^## ", line)) {
              # Heading 2
              doc <- officer::body_add_par(doc, gsub("^## ", "", line), 
                                          style = "heading 2")
            } else if(grepl("^### ", line)) {
              # Heading 3
              doc <- officer::body_add_par(doc, gsub("^### ", "", line), 
                                          style = "heading 3")
            } else if(nchar(trimws(line)) > 0) {
              # Regular paragraph - handle bold and italic
              clean_line <- line
              clean_line <- gsub("\\*\\*(.+?)\\*\\*", "\\1", clean_line)
              clean_line <- gsub("\\*(.+?)\\*", "\\1", clean_line)
              doc <- officer::body_add_par(doc, clean_line)
            }
          }
          
          print(doc, target = file)
        } else {
          # Fallback to plain text
          writeLines(gsub("\\*", "", values$writeup_text), file)
          showNotification("Note: 'officer' package not available. Saved as plain text.", 
                          type = "warning")
        }
      } else {
        # Plain text or markdown
        text_to_write <- if(input$writeup_format == "txt") {
          gsub("\\*", "", gsub("## |### ", "", values$writeup_text))
        } else {
          values$writeup_text
        }
        writeLines(text_to_write, file)
      }
    },
    contentType = function() {
      switch(input$writeup_format,
        "txt" = "text/plain",
        "md" = "text/markdown",
        "docx" = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      )
    }
  )

  # =================== DOWNLOADS =============================================
  
  output$download_filtered <- downloadHandler(
    filename = function() paste0("filtered_data_", format(Sys.Date(), "%Y-%m-%d"), ".csv"),
    content = function(file) write.csv(values$filtered_data, file, row.names = FALSE),
    contentType = "text/csv"
  )
  
  output$download_es <- downloadHandler(
    filename = function() paste0("effect_sizes_", format(Sys.Date(), "%Y-%m-%d"), ".csv"),
    content = function(file) write.csv(values$es_data, file, row.names = FALSE),
    contentType = "text/csv"
  )
  
  output$download_results <- downloadHandler(
    filename = function() paste0("results_", format(Sys.Date(), "%Y-%m-%d"), ".txt"),
    content = function(file) {
      sink(file)
      
      if(!is.null(values$ma_model)) {
        cat("=== META-ANALYSIS RESULTS ===\n\n")
        print(summary(values$ma_model))
        
        # Add Egger's test results if run
        if(isTRUE(input$run_egger)) {
          cat("\n\n=== EGGER'S TEST FOR FUNNEL PLOT ASYMMETRY ===\n\n")
        # Handle both rma and rma.mv models for Egger's test
        if(inherits(values$ma_model, "rma.mv")) {
          agg_data <- values$es_data %>%
            group_by(StudyGroup) %>%
            summarise(
              yi_agg = mean(yi, na.rm = TRUE),
              vi_agg = mean(vi, na.rm = TRUE),
              .groups = 'drop'
            )
          
          temp_model <- tryCatch(
            rma(yi = yi_agg, vi = vi_agg, data = agg_data, method = "REML"),
            error = function(e) NULL
          )
          
          if(!is.null(temp_model)) {
            egger <- regtest(temp_model)
          } else {
            egger <- NULL
          }
        } else {
          egger <- regtest(values$ma_model)
        }
        
          cat("z =", round(egger$zval, 3), "\n")
          cat("p =", format.pval(egger$pval, 3), "\n")
          
          if(egger$pval < 0.05) {
            cat("\n*** Significant asymmetry detected (p < 0.05) ***\n")
          } else {
            cat("\nNo significant asymmetry detected.\n")
          }
        }
        
        # Add trim-and-fill results if available
        if(!is.null(values$trimfill_model) && isTRUE(input$run_trimfill)) {
          cat("\n\n=== DUVAL & TWEEDIE'S TRIM-AND-FILL PROCEDURE ===\n\n")
          print(values$trimfill_model)
          
          cat("\n--- Comparison ---\n")
          cat("Original estimate:", round(coef(values$ma_model), 4), "\n")
          cat("Original 95% CI: [", round(values$ma_model$ci.lb, 4), ", ", 
              round(values$ma_model$ci.ub, 4), "]\n")
          cat("\nAdjusted estimate:", round(coef(values$trimfill_model), 4), "\n")
          cat("Adjusted 95% CI: [", round(values$trimfill_model$ci.lb, 4), ", ", 
              round(values$trimfill_model$ci.ub, 4), "]\n")
          cat("\nNumber of imputed studies:", values$trimfill_model$k0, "\n")
          cat("Side:", input$trimfill_side, "\n")
        }
      }
      
      sink()
    },
    contentType = "text/plain"
  )
  
  output$download_enhanced_forest <- downloadHandler(
    filename = function() {
      ext <- if(input$plot_format == "pdf") "pdf" else "png"
      paste0("enhanced_forest_", format(Sys.Date(), "%Y-%m-%d"), ".", ext)
    },
    content = function(file) {
      p <- forest_plot()
      req(p)
      
      dpi <- if (input$plot_format == "png" && !is.null(input$export_dpi)) {
        input$export_dpi
      } else {
        300
      }
      
      if (input$plot_format == "pdf") {
        ggsave(
          filename = file,
          plot     = p,
          device   = cairo_pdf,
          width    = input$forest_width,
          height   = input$forest_height,
          units    = "in"
        )
      } else {
        ggsave(
          filename = file,
          plot     = p,
          device   = "png",
          width    = input$forest_width,
          height   = input$forest_height,
          units    = "in",
          dpi      = dpi
        )
      }
    },
    contentType = "image/png"  # Default to PNG
  )
  
  output$download_forest <- downloadHandler(
    filename = function() {
      ext <- if(input$plot_format == "pdf") "pdf" else "png"
      paste0("enhanced_forest_", format(Sys.Date(), "%Y-%m-%d"), ".", ext)
    },
    content = function(file) {
      # Use the same enhanced forest plot as the main download button
      p <- forest_plot()
      req(p)
      
      dpi <- if (input$plot_format == "png" && !is.null(input$export_dpi)) {
        input$export_dpi
      } else {
        300
      }
      
      if (input$plot_format == "pdf") {
        ggsave(
          filename = file,
          plot     = p,
          device   = cairo_pdf,
          width    = input$forest_width,
          height   = input$forest_height,
          units    = "in"
        )
      } else {
        ggsave(
          filename = file,
          plot     = p,
          device   = "png",
          width    = input$forest_width,
          height   = input$forest_height,
          units    = "in",
          dpi      = dpi
        )
      }
    },
    contentType = "image/png"  # Add missing contentType!
  )
  
  output$download_funnel <- downloadHandler(
    filename = function() {
      ext <- if(input$plot_format == "pdf") "pdf" else "png"
      paste0("funnel_", format(Sys.Date(), "%Y-%m-%d"), ".", ext)
    },
    content = function(file) {
      # Create temp file with correct extension
      ext <- if(input$plot_format == "pdf") ".pdf" else ".png"
      temp_file <- tempfile(fileext = ext)
      
      dpi <- if(input$plot_format == "png" && !is.null(input$export_dpi)) input$export_dpi else 300
      
      if(input$plot_format == "pdf") {
        cairo_pdf(temp_file, width = 8, height = 8)
      } else {
        png(temp_file, width = 8 * dpi, height = 8 * dpi, res = dpi, type = "cairo")
      }
      
      if(!is.null(values$ma_model)) {
        if(input$contour_enhanced) {
          funnel(values$ma_model, level = c(90, 95, 99), 
                shade = c("white", "gray75", "gray60"), legend = TRUE)
        } else {
          funnel(values$ma_model)
        }
        
        if(input$label_funnel == TRUE && nrow(values$es_data) <= input$max_funnel_labels) {
          text(values$es_data$yi, sqrt(values$es_data$vi),
              values$es_data$StudyID, cex = input$label_size, pos = 3)
        }
      }
      
      dev.off()
      
      # Copy temp file to download location
      file.copy(temp_file, file, overwrite = TRUE)
      unlink(temp_file)
    },
    contentType = "image/png"  # Add missing contentType!
  )
  
  output$download_trimfill_funnel <- downloadHandler(
    filename = function() {
      ext <- if(input$plot_format == "pdf") "pdf" else "png"
      paste0("trimfill_funnel_", format(Sys.Date(), "%Y-%m-%d"), ".", ext)
    },
    content = function(file) {
      req(values$trimfill_model)
      
      dpi <- if(input$plot_format == "png" && !is.null(input$export_dpi)) input$export_dpi else 300
      
      # Write directly to file parameter
      if(input$plot_format == "pdf") {
        cairo_pdf(file, width = 8, height = 8)
      } else {
        png(file, width = 8 * dpi, height = 8 * dpi, res = dpi, type = "cairo")
      }
      
      # Match the displayed plot - include significance contours
      funnel(values$trimfill_model,
             level = c(90, 95, 99),
             shade = c("white", "gray55", "gray75"),
             refline = 0,
             xlab = "Effect Size",
             main = "Trim-and-Fill Adjusted Funnel Plot")
      
      # Add note if based on aggregated data
      if(inherits(values$ma_model, "rma.mv")) {
        mtext("Note: Based on aggregated effect sizes (one per study-group)", 
              side = 1, line = 4, cex = 0.8, col = "blue")
      }
      
      dev.off()
    },
    contentType = "image/png"
  )
  
}

# Run app
shinyApp(ui = ui, server = server)
