#!/usr/bin/env python3
"""
Systematic Review Screening Tool
A GUI application for searching, retrieving, and screening academic articles
from PubMed and Cochrane for systematic reviews.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import pandas as pd
import requests
import re
import csv
import os
from datetime import datetime
import threading
import time
from urllib.parse import urlencode
from xml.etree import ElementTree as ET
try:
    from Bio import Entrez
    BIOPYTHON_AVAILABLE = True
except ImportError:
    BIOPYTHON_AVAILABLE = False
    print("Warning: Biopython not available. PubMed search will use alternative method.")

class SystematicReviewTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Systematic Review Screening Tool")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        # Data storage
        self.all_articles = []
        self.current_index = 0
        self.included_articles = []
        self.excluded_articles = []
        self.inclusion_keywords = []
        
        # Email for PubMed API (required by NCBI)
        self.email = "your.email@example.com"  # User should change this
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the main user interface"""
        # Create notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab 1: Search
        self.search_frame = ttk.Frame(notebook)
        notebook.add(self.search_frame, text="Search & Retrieve")
        self.setup_search_tab()
        
        # Tab 2: Screening
        self.screening_frame = ttk.Frame(notebook)
        notebook.add(self.screening_frame, text="Screen Articles")
        self.setup_screening_tab()
        
        # Tab 3: Results
        self.results_frame = ttk.Frame(notebook)
        notebook.add(self.results_frame, text="Results")
        self.setup_results_tab()
        
    def setup_search_tab(self):
        """Setup the search and retrieval tab"""
        # Email configuration
        email_frame = ttk.LabelFrame(self.search_frame, text="Configuration", padding=10)
        email_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(email_frame, text="Email (required for PubMed):").pack(anchor='w')
        self.email_var = tk.StringVar(value=self.email)
        email_entry = ttk.Entry(email_frame, textvariable=self.email_var, width=50)
        email_entry.pack(fill='x', pady=2)
        
        # Search terms
        search_frame = ttk.LabelFrame(self.search_frame, text="Search Parameters", padding=10)
        search_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(search_frame, text="Search Terms:").pack(anchor='w')
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=50)
        search_entry.pack(fill='x', pady=2)
        
        ttk.Label(search_frame, text="Max Results:").pack(anchor='w', pady=(10,0))
        self.max_results_var = tk.StringVar(value="100")
        max_results_entry = ttk.Entry(search_frame, textvariable=self.max_results_var, width=10)
        max_results_entry.pack(anchor='w', pady=2)
        
        # Database selection
        db_frame = ttk.LabelFrame(self.search_frame, text="Databases", padding=10)
        db_frame.pack(fill='x', padx=10, pady=5)
        
        self.pubmed_var = tk.BooleanVar(value=True)
        self.cochrane_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(db_frame, text="PubMed", variable=self.pubmed_var).pack(anchor='w')
        ttk.Checkbutton(db_frame, text="Cochrane (manual CSV import)", variable=self.cochrane_var).pack(anchor='w')
        
        # Cochrane file selection
        cochrane_file_frame = ttk.Frame(db_frame)
        cochrane_file_frame.pack(fill='x', pady=5)
        
        ttk.Label(cochrane_file_frame, text="Cochrane CSV file:").pack(anchor='w')
        file_select_frame = ttk.Frame(cochrane_file_frame)
        file_select_frame.pack(fill='x')
        
        self.cochrane_file_var = tk.StringVar()
        ttk.Entry(file_select_frame, textvariable=self.cochrane_file_var, state='readonly').pack(side='left', fill='x', expand=True)
        ttk.Button(file_select_frame, text="Browse", command=self.select_cochrane_file).pack(side='right', padx=(5,0))
        
        # Search button and progress
        button_frame = ttk.Frame(self.search_frame)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        self.search_button = ttk.Button(button_frame, text="Start Search", command=self.start_search)
        self.search_button.pack(side='left')
        
        self.progress_var = tk.StringVar(value="Ready to search")
        ttk.Label(button_frame, textvariable=self.progress_var).pack(side='left', padx=(10,0))
        
        # Results preview
        results_frame = ttk.LabelFrame(self.search_frame, text="Search Results Preview", padding=10)
        results_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.results_text = scrolledtext.ScrolledText(results_frame, height=15)
        self.results_text.pack(fill='both', expand=True)
        
    def setup_screening_tab(self):
        """Setup the article screening tab"""
        # Inclusion keywords
        keywords_frame = ttk.LabelFrame(self.screening_frame, text="Inclusion Keywords", padding=10)
        keywords_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(keywords_frame, text="Keywords (comma-separated):").pack(anchor='w')
        self.keywords_var = tk.StringVar()
        keywords_entry = ttk.Entry(keywords_frame, textvariable=self.keywords_var, width=50)
        keywords_entry.pack(fill='x', pady=2)
        
        ttk.Button(keywords_frame, text="Update Keywords", command=self.update_keywords).pack(anchor='w', pady=5)
        
        # Article display
        article_frame = ttk.LabelFrame(self.screening_frame, text="Article Screening", padding=10)
        article_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Progress info
        self.screening_progress_var = tk.StringVar(value="No articles loaded")
        ttk.Label(article_frame, textvariable=self.screening_progress_var, font=('Arial', 10, 'bold')).pack(pady=5)
        
        # Article content
        content_frame = ttk.Frame(article_frame)
        content_frame.pack(fill='both', expand=True)
        
        # Article display with scrollbar
        self.article_display = scrolledtext.ScrolledText(content_frame, height=20, wrap='word')
        self.article_display.pack(fill='both', expand=True)
        
        # Configure text tags for highlighting
        self.article_display.tag_configure("highlight", background="lightgreen", foreground="black")
        self.article_display.tag_configure("title", font=('Arial', 14, 'bold'))
        self.article_display.tag_configure("authors", font=('Arial', 10, 'italic'))
        self.article_display.tag_configure("abstract", font=('Arial', 10))
        
        # Decision buttons
        decision_frame = ttk.Frame(article_frame)
        decision_frame.pack(fill='x', pady=10)
        
        ttk.Button(decision_frame, text="◀ Previous", command=self.previous_article).pack(side='left')
        
        button_spacer = ttk.Frame(decision_frame)
        button_spacer.pack(side='left', fill='x', expand=True)
        
        self.exclude_button = ttk.Button(button_spacer, text="❌ Exclude", command=self.exclude_article, style='Exclude.TButton')
        self.exclude_button.pack(side='left', padx=20)
        
        self.include_button = ttk.Button(button_spacer, text="✅ Include", command=self.include_article, style='Include.TButton')
        self.include_button.pack(side='right', padx=20)
        
        ttk.Button(decision_frame, text="Next ▶", command=self.next_article).pack(side='right')
        
        # Style the buttons
        style = ttk.Style()
        style.configure('Include.TButton', foreground='green')
        style.configure('Exclude.TButton', foreground='red')
        
    def setup_results_tab(self):
        """Setup the results and export tab"""
        # Summary statistics
        stats_frame = ttk.LabelFrame(self.results_frame, text="Screening Summary", padding=10)
        stats_frame.pack(fill='x', padx=10, pady=5)
        
        self.stats_var = tk.StringVar(value="No screening completed yet")
        ttk.Label(stats_frame, textvariable=self.stats_var, font=('Arial', 12)).pack()
        
        # Export options
        export_frame = ttk.LabelFrame(self.results_frame, text="Export Results", padding=10)
        export_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(export_frame, text="Export All Results", command=self.export_all_results).pack(pady=5)
        ttk.Button(export_frame, text="Export Included Only", command=self.export_included).pack(pady=5)
        ttk.Button(export_frame, text="Export Excluded Only", command=self.export_excluded).pack(pady=5)
        
        # Results preview
        preview_frame = ttk.LabelFrame(self.results_frame, text="Results Preview", padding=10)
        preview_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.results_preview = scrolledtext.ScrolledText(preview_frame, height=15)
        self.results_preview.pack(fill='both', expand=True)
        
    def select_cochrane_file(self):
        """Select Cochrane CSV file"""
        filename = filedialog.askopenfilename(
            title="Select Cochrane CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.cochrane_file_var.set(filename)
    
    def start_search(self):
        """Start the search process in a separate thread"""
        self.email = self.email_var.get().strip()
        search_terms = self.search_var.get().strip()
        
        if not search_terms:
            messagebox.showerror("Error", "Please enter search terms")
            return
        
        if not self.email:
            messagebox.showerror("Error", "Please enter your email address")
            return
        
        try:
            max_results = int(self.max_results_var.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for max results")
            return
        
        # Disable search button and start search in thread
        self.search_button.config(state='disabled')
        self.progress_var.set("Searching...")
        
        search_thread = threading.Thread(target=self.perform_search, 
                                        args=(search_terms, max_results))
        search_thread.daemon = True
        search_thread.start()
    
    def perform_search(self, search_terms, max_results):
        """Perform the actual search"""
        try:
            self.all_articles = []
            
            # Search PubMed
            if self.pubmed_var.get():
                self.progress_var.set("Searching PubMed...")
                pubmed_articles = self.search_pubmed(search_terms, max_results)
                self.all_articles.extend(pubmed_articles)
                self.progress_var.set(f"Found {len(pubmed_articles)} PubMed articles")
                time.sleep(0.5)
            
            # Load Cochrane CSV
            if self.cochrane_var.get() and self.cochrane_file_var.get():
                self.progress_var.set("Loading Cochrane articles...")
                cochrane_articles = self.load_cochrane_csv(self.cochrane_file_var.get())
                self.all_articles.extend(cochrane_articles)
                self.progress_var.set(f"Loaded {len(cochrane_articles)} Cochrane articles")
                time.sleep(0.5)
            
            # Save to CSV
            if self.all_articles:
                self.save_search_results()
                self.display_search_results()
                self.progress_var.set(f"Search completed! Found {len(self.all_articles)} total articles")
                
                # Initialize screening
                self.current_index = 0
                self.update_screening_display()
            else:
                self.progress_var.set("No articles found")
                
        except Exception as e:
            self.progress_var.set(f"Error: {str(e)}")
            messagebox.showerror("Search Error", f"An error occurred during search: {str(e)}")
        finally:
            self.search_button.config(state='normal')
    
    def search_pubmed(self, search_terms, max_results):
        """Search PubMed using Biopython or direct API"""
        articles = []
        
        if BIOPYTHON_AVAILABLE:
            try:
                Entrez.email = self.email
                
                # Search for article IDs
                handle = Entrez.esearch(db="pubmed", term=search_terms, retmax=max_results)
                search_results = Entrez.read(handle)
                handle.close()
                
                id_list = search_results["IdList"]
                
                if id_list:
                    # Fetch article details
                    handle = Entrez.efetch(db="pubmed", id=id_list, rettype="xml", retmode="xml")
                    records = Entrez.read(handle)
                    handle.close()
                    
                    for i, record in enumerate(records['PubmedArticle']):
                        try:
                            article = self.parse_pubmed_record(record, i + 1)
                            articles.append(article)
                        except Exception as e:
                            print(f"Error parsing record {i}: {e}")
                            continue
                            
            except Exception as e:
                print(f"Biopython search failed: {e}")
                # Fallback to direct API
                articles = self.search_pubmed_direct(search_terms, max_results)
        else:
            articles = self.search_pubmed_direct(search_terms, max_results)
        
        return articles
    
    def search_pubmed_direct(self, search_terms, max_results):
        """Direct PubMed API search without Biopython"""
        articles = []
        
        try:
            # Search for PMIDs
            search_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
            search_params = {
                'db': 'pubmed',
                'term': search_terms,
                'retmax': max_results,
                'retmode': 'json',
                'email': self.email
            }
            
            response = requests.get(search_url, params=search_params, timeout=30)
            response.raise_for_status()
            search_data = response.json()
            
            pmids = search_data.get('esearchresult', {}).get('idlist', [])
            
            if pmids:
                # Fetch article details in batches
                fetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
                
                batch_size = 100
                for i in range(0, len(pmids), batch_size):
                    batch_pmids = pmids[i:i + batch_size]
                    
                    fetch_params = {
                        'db': 'pubmed',
                        'id': ','.join(batch_pmids),
                        'rettype': 'xml',
                        'retmode': 'xml',
                        'email': self.email
                    }
                    
                    response = requests.get(fetch_url, params=fetch_params, timeout=60)
                    response.raise_for_status()
                    
                    # Parse XML
                    root = ET.fromstring(response.text)
                    
                    for j, pubmed_article in enumerate(root.findall('.//PubmedArticle')):
                        try:
                            article = self.parse_pubmed_xml(pubmed_article, len(articles) + 1)
                            articles.append(article)
                        except Exception as e:
                            print(f"Error parsing XML record {j}: {e}")
                            continue
                            
        except Exception as e:
            print(f"Direct PubMed search failed: {e}")
        
        return articles
    
    def parse_pubmed_record(self, record, sr_no):
        """Parse PubMed record from Biopython"""
        try:
            medline_citation = record['MedlineCitation']
            article = medline_citation['Article']
            
            title = article.get('ArticleTitle', 'No title')
            pmid = str(medline_citation.get('PMID', ''))
            abstract_list = article.get('Abstract', {}).get('AbstractText', [])
            
            abstract = ""
            if abstract_list:
                if isinstance(abstract_list, list):
                    abstract = ' '.join([str(ab) for ab in abstract_list])
                else:
                    abstract = str(abstract_list)
            
            # Parse authors
            author_list = article.get('AuthorList', [])
            authors = []
            for author in author_list:
                if 'LastName' in author and 'ForeName' in author:
                    authors.append(f"{author['LastName']} {author['ForeName']}")
                elif 'CollectiveName' in author:
                    authors.append(author['CollectiveName'])
            
            authors_str = '; '.join(authors) if authors else 'No authors listed'
            
            return {
                'Sr_No': sr_no,
                'Title': title,
                'ID_Link': f"PMID: {pmid}",
                'Abstract': abstract,
                'Authors': authors_str,
                'Source': 'PubMed',
                'Status': 'Pending'
            }
        except Exception as e:
            print(f"Error parsing Biopython record: {e}")
            return None
    
    def parse_pubmed_xml(self, pubmed_article, sr_no):
        """Parse PubMed XML directly"""
        try:
            # Extract basic information
            article = pubmed_article.find('.//Article')
            
            title_elem = article.find('ArticleTitle')
            title = title_elem.text if title_elem is not None else 'No title'
            
            pmid_elem = pubmed_article.find('.//PMID')
            pmid = pmid_elem.text if pmid_elem is not None else ''
            
            # Extract abstract
            abstract_text = ""
            abstract_elem = article.find('.//Abstract')
            if abstract_elem is not None:
                abstract_parts = []
                for text_elem in abstract_elem.findall('.//AbstractText'):
                    if text_elem.text:
                        abstract_parts.append(text_elem.text)
                abstract_text = ' '.join(abstract_parts)
            
            # Extract authors
            authors = []
            author_list = article.find('.//AuthorList')
            if author_list is not None:
                for author in author_list.findall('Author'):
                    last_name = author.find('LastName')
                    fore_name = author.find('ForeName')
                    collective_name = author.find('CollectiveName')
                    
                    if last_name is not None and fore_name is not None:
                        authors.append(f"{last_name.text} {fore_name.text}")
                    elif collective_name is not None:
                        authors.append(collective_name.text)
            
            authors_str = '; '.join(authors) if authors else 'No authors listed'
            
            return {
                'Sr_No': sr_no,
                'Title': title,
                'ID_Link': f"PMID: {pmid}",
                'Abstract': abstract_text,
'Authors': authors_str,
                'Source': 'PubMed',
                'Status': 'Pending'
            }
        except Exception as e:
            print(f"Error parsing Biopython record: {e}")
            return None
    
    def parse_pubmed_xml(self, pubmed_article, sr_no):
        """Parse PubMed XML directly"""
        try:
            # Extract basic information
            article = pubmed_article.find('.//Article')
            
            title_elem = article.find('ArticleTitle')
            title = title_elem.text if title_elem is not None else 'No title'
            
            pmid_elem = pubmed_article.find('.//PMID')
            pmid = pmid_elem.text if pmid_elem is not None else ''
            
            # Extract abstract
            abstract_text = ""
            abstract_elem = article.find('.//Abstract')
            if abstract_elem is not None:
                abstract_parts = []
                for text_elem in abstract_elem.findall('.//AbstractText'):
                    if text_elem.text:
                        abstract_parts.append(text_elem.text)
                abstract_text = ' '.join(abstract_parts)
            
            # Extract authors
            authors = []
            author_list = article.find('.//AuthorList')
            if author_list is not None:
                for author in author_list.findall('Author'):
                    last_name = author.find('LastName')
                    fore_name = author.find('ForeName')
                    collective_name = author.find('CollectiveName')
                    
                    if last_name is not None and fore_name is not None:
                        authors.append(f"{last_name.text} {fore_name.text}")
                    elif collective_name is not None:
                        authors.append(collective_name.text)
            
            authors_str = '; '.join(authors) if authors else 'No authors listed'
            
            return {
                'Sr_No': sr_no,
                'Title': title,
                'ID_Link': f"PMID: {pmid}",
                'Abstract': abstract_text,
                'Authors': authors_str,
                'Source': 'PubMed',
                'Status': 'Pending'
            }
        except Exception as e:
            print(f"Error parsing XML: {e}")
            return None
    
    def load_cochrane_csv(self, filename):
        """Load articles from Cochrane CSV file"""
        articles = []
        
        try:
            df = pd.read_csv(filename)
            
            # Try to map common Cochrane CSV column names
            column_mapping = {
                'title': ['Title', 'title', 'TITLE'],
                'authors': ['Authors', 'authors', 'AUTHORS', 'Author'],
                'abstract': ['Abstract', 'abstract', 'ABSTRACT'],
                'doi': ['DOI', 'doi', 'ID', 'id']
            }
            
            actual_columns = {}
            for key, possible_names in column_mapping.items():
                for name in possible_names:
                    if name in df.columns:
                        actual_columns[key] = name
                        break
            
            for i, row in df.iterrows():
                try:
                    title = row.get(actual_columns.get('title', ''), 'No title')
                    authors = row.get(actual_columns.get('authors', ''), 'No authors listed')
                    abstract = row.get(actual_columns.get('abstract', ''), '')
                    doi = row.get(actual_columns.get('doi', ''), '')
                    
                    article = {
                        'Sr_No': len(self.all_articles) + i + 1,
                        'Title': str(title),
                        'ID_Link': f"DOI: {doi}" if doi else 'No ID',
                        'Abstract': str(abstract),
                        'Authors': str(authors),
                        'Source': 'Cochrane',
                        'Status': 'Pending'
                    }
                    articles.append(article)
                except Exception as e:
                    print(f"Error processing Cochrane row {i}: {e}")
                    continue
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Cochrane CSV: {e}")
        
        return articles
    
    def save_search_results(self):
        """Save all search results to CSV"""
        if self.all_articles:
            df = pd.DataFrame(self.all_articles)
            df.to_csv('search_results.csv', index=False)
    
    def display_search_results(self):
        """Display search results in the text widget"""
        self.results_text.delete(1.0, tk.END)
        
        summary = f"Search Results Summary:\n"
        summary += f"Total articles found: {len(self.all_articles)}\n"
        
        pubmed_count = sum(1 for article in self.all_articles if article['Source'] == 'PubMed')
        cochrane_count = sum(1 for article in self.all_articles if article['Source'] == 'Cochrane')
        
        summary += f"PubMed articles: {pubmed_count}\n"
        summary += f"Cochrane articles: {cochrane_count}\n\n"
        
        self.results_text.insert(tk.END, summary)
        
        # Display first few articles
        for i, article in enumerate(self.all_articles[:5]):
            self.results_text.insert(tk.END, f"{i+1}. {article['Title']}\n")
            self.results_text.insert(tk.END, f"   Authors: {article['Authors']}\n")
            self.results_text.insert(tk.END, f"   {article['ID_Link']}\n")
            self.results_text.insert(tk.END, f"   Source: {article['Source']}\n\n")
        
        if len(self.all_articles) > 5:
            self.results_text.insert(tk.END, f"... and {len(self.all_articles) - 5} more articles\n")
    
    def update_keywords(self):
        """Update inclusion keywords"""
        keywords_text = self.keywords_var.get().strip()
        if keywords_text:
            self.inclusion_keywords = [kw.strip().lower() for kw in keywords_text.split(',') if kw.strip()]
            messagebox.showinfo("Keywords Updated", f"Updated {len(self.inclusion_keywords)} inclusion keywords")
        else:
            self.inclusion_keywords = []
        
        # Refresh current article display
        if self.all_articles and 0 <= self.current_index < len(self.all_articles):
            self.display_current_article()
    
    def update_screening_display(self):
        """Update the screening progress display"""
        if not self.all_articles:
            self.screening_progress_var.set("No articles loaded")
            return
        
        total = len(self.all_articles)
        current = self.current_index + 1
        included = len(self.included_articles)
        excluded = len(self.excluded_articles)
        
        self.screening_progress_var.set(
            f"Article {current} of {total} | Included: {included} | Excluded: {excluded}"
        )
        
        self.display_current_article()
        self.update_results_summary()
    
    def display_current_article(self):
        """Display the current article with keyword highlighting"""
        if not self.all_articles or self.current_index >= len(self.all_articles):
            self.article_display.delete(1.0, tk.END)
            self.article_display.insert(tk.END, "No more articles to review")
            return
        
        article = self.all_articles[self.current_index]
        
        # Clear display
        self.article_display.delete(1.0, tk.END)
        
        # Display title
        title_text = f"TITLE: {article['Title']}\n\n"
        self.article_display.insert(tk.END, title_text, "title")
        
        # Display authors
        authors_text = f"AUTHORS: {article['Authors']}\n\n"
        self.article_display.insert(tk.END, authors_text, "authors")
        
        # Display ID/Link
        id_text = f"ID: {article['ID_Link']}\n\n"
        self.article_display.insert(tk.END, id_text)
        
        # Display abstract with highlighting
        abstract_text = f"ABSTRACT:\n{article['Abstract']}\n\n"
        self.insert_text_with_highlights(abstract_text, "abstract")
        
        # Display source
        source_text = f"SOURCE: {article['Source']}\n"
        self.article_display.insert(tk.END, source_text)
    
    def insert_text_with_highlights(self, text, tag):
        """Insert text with keyword highlighting"""
        if not self.inclusion_keywords:
            self.article_display.insert(tk.END, text, tag)
            return
        
        # Create pattern for all keywords
        patterns = []
        for keyword in self.inclusion_keywords:
            patterns.append(re.escape(keyword))
        
        if not patterns:
            self.article_display.insert(tk.END, text, tag)
            return
        
        pattern = '|'.join(patterns)
        
        # Split text by matches
        parts = re.split(f'({pattern})', text, flags=re.IGNORECASE)
        
        for part in parts:
            if part.lower() in [kw.lower() for kw in self.inclusion_keywords]:
                self.article_display.insert(tk.END, part, "highlight")
            else:
                self.article_display.insert(tk.END, part, tag)
    
    def include_article(self):
        """Include current article"""
        if self.all_articles and 0 <= self.current_index < len(self.all_articles):
            article = self.all_articles[self.current_index].copy()
            article['Status'] = 'Included'
            article['Decision_Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            self.included_articles.append(article)
            self.all_articles[self.current_index]['Status'] = 'Included'
            
            self.next_article()
    
    def exclude_article(self):
        """Exclude current article"""
        if self.all_articles and 0 <= self.current_index < len(self.all_articles):
            article = self.all_articles[self.current_index].copy()
            article['Status'] = 'Excluded'
            article['Decision_Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            self.excluded_articles.append(article)
            self.all_articles[self.current_index]['Status'] = 'Excluded'
            
            self.next_article()
    
    def next_article(self):
        """Move to next article"""
        if self.current_index < len(self.all_articles) - 1:
            self.current_index += 1
            self.update_screening_display()
        else:
            messagebox.showinfo("Screening Complete", 
                              "You have reviewed all articles! Check the Results tab.")
    
    def previous_article(self):
        """Move to previous article"""
        if self.current_index > 0:
            self.current_index -= 1
            self.update_screening_display()
    
    def update_results_summary(self):
        """Update the results summary"""
        total = len(self.all_articles)
        included = len(self.included_articles)
        excluded = len(self.excluded_articles)
        pending = total - included - excluded
        
        summary = f"Total Articles: {total} | Included: {included} | Excluded: {excluded} | Pending: {pending}"
        self.stats_var.set(summary)
        
        # Update results preview
        self.update_results_preview()
    
    def update_results_preview(self):
        """Update the results preview text"""
        self.results_preview.delete(1.0, tk.END)
        
        # Summary statistics
        total = len(self.all_articles)
        included = len(self.included_articles)
        excluded = len(self.excluded_articles)
        pending = total - included - excluded
        
        summary_text = f"""SCREENING SUMMARY
=================
Total Articles Found: {total}
Included Articles: {included}
Excluded Articles: {excluded}
Pending Review: {pending}

"""
        self.results_preview.insert(tk.END, summary_text)
        
        # Show included articles
        if self.included_articles:
            self.results_preview.insert(tk.END, "INCLUDED ARTICLES:\n")
            self.results_preview.insert(tk.END, "-" * 50 + "\n")
            for i, article in enumerate(self.included_articles, 1):
                self.results_preview.insert(tk.END, f"{i}. {article['Title']}\n")
                self.results_preview.insert(tk.END, f"   Authors: {article['Authors']}\n")
                self.results_preview.insert(tk.END, f"   {article['ID_Link']}\n")
                self.results_preview.insert(tk.END, f"   Decision Date: {article.get('Decision_Date', 'N/A')}\n\n")
        
        # Show some excluded articles
        if self.excluded_articles:
            self.results_preview.insert(tk.END, "\nEXCLUDED ARTICLES (first 5):\n")
            self.results_preview.insert(tk.END, "-" * 50 + "\n")
            for i, article in enumerate(self.excluded_articles[:5], 1):
                self.results_preview.insert(tk.END, f"{i}. {article['Title']}\n")
                self.results_preview.insert(tk.END, f"   Authors: {article['Authors']}\n")
                self.results_preview.insert(tk.END, f"   Decision Date: {article.get('Decision_Date', 'N/A')}\n\n")
            
            if len(self.excluded_articles) > 5:
                self.results_preview.insert(tk.END, f"... and {len(self.excluded_articles) - 5} more excluded articles\n")
    
    def export_all_results(self):
        """Export all results to CSV"""
        if not self.all_articles:
            messagebox.showwarning("No Data", "No articles to export")
            return
        
        try:
            # Update all articles with current status
            for article in self.all_articles:
                if article['Status'] == 'Included':
                    # Find the included version with decision date
                    for inc_article in self.included_articles:
                        if inc_article['Sr_No'] == article['Sr_No']:
                            article['Decision_Date'] = inc_article.get('Decision_Date', '')
                            break
                elif article['Status'] == 'Excluded':
                    # Find the excluded version with decision date
                    for exc_article in self.excluded_articles:
                        if exc_article['Sr_No'] == article['Sr_No']:
                            article['Decision_Date'] = exc_article.get('Decision_Date', '')
                            break
            
            df = pd.DataFrame(self.all_articles)
            filename = f"all_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            df.to_csv(filename, index=False)
            
            messagebox.showinfo("Export Successful", f"All results exported to {filename}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Error exporting results: {e}")
    
    def export_included(self):
        """Export included articles to CSV"""
        if not self.included_articles:
            messagebox.showwarning("No Data", "No included articles to export")
            return
        
        try:
            df = pd.DataFrame(self.included_articles)
            filename = f"included_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            df.to_csv(filename, index=False)
            
            messagebox.showinfo("Export Successful", f"Included articles exported to {filename}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Error exporting included articles: {e}")
    
    def export_excluded(self):
        """Export excluded articles to CSV"""
        if not self.excluded_articles:
            messagebox.showwarning("No Data", "No excluded articles to export")
            return
        
        try:
            df = pd.DataFrame(self.excluded_articles)
            filename = f"excluded_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            df.to_csv(filename, index=False)
            
            messagebox.showinfo("Export Successful", f"Excluded articles exported to {filename}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Error exporting excluded articles: {e}")


def main():
    """Main function to run the application"""
    root = tk.Tk()
    
    # Set application icon (if available)
    try:
        root.iconbitmap('icon.ico')  # You can add an icon file
    except:
        pass
    
    # Configure ttk theme
    style = ttk.Style()
    available_themes = style.theme_names()
    
    # Use a modern theme if available
    if 'clam' in available_themes:
        style.theme_use('clam')
    elif 'alt' in available_themes:
        style.theme_use('alt')
    
    # Create and run the application
    app = SystematicReviewTool(root)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()
