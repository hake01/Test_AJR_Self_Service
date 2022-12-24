package com.sn.ajr;

public class UpdateSlidePojo {

	private String pageObjectId;
	
	private int slideNo;
	
	private String oldText;
	
	private String newText;

	public UpdateSlidePojo(int slideNo, String oldText, String newText) {
		super();
		this.slideNo = slideNo;
		this.oldText = oldText;
		this.newText = newText;
	}

	public String getPageObjectId() {
		return pageObjectId;
	}

	public void setPageObjectId(String pageObjectId) {
		this.pageObjectId = pageObjectId;
	}

	public String getOldText() {
		return oldText;
	}

	public void setOldText(String oldText) {
		this.oldText = oldText;
	}

	public String getNewText() {
		return newText;
	}

	public void setNewText(String newText) {
		this.newText = newText;
	}

	public int getSlideNo() {
		return slideNo;
	}

	public void setSlideNo(int slideNo) {
		this.slideNo = slideNo;
	}
	
	
}
